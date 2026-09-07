"""
Build script: raw NCRB Excel files -> category JSON files for the crime atlas.

Usage:
    python3 build_data.py

Reads: All2014.xlsx ... All2019.xlsx, ncrb_district_crosswalk.csv, districts.geojson
Writes: ../data/<category>.json for each category, plus categories.json
"""
import pandas as pd
import json
import re
import os

YEARS = [2014, 2015, 2016, 2017, 2018, 2019]
SIMPLE_FORMAT_YEARS = {2014, 2015, 2016}   # single-row headers
NESTED_FORMAT_YEARS = {2017, 2018, 2019}   # multi-row nested headers

SHEET_MAP_SIMPLE = {'IPC':'Sheet1','SLL':'Sheet2','Women':'Sheet3','Children':'Sheet4','SC':'Sheet5','ST':'Sheet6'}
SHEET_MAP_NESTED = {'IPC':'Sheet1','SLL':'Sheet2','Women':'Sheet3','Children':'Sheet4','SC':'Sheet5',
                     'ST':'Sheet6','Cyber':'Sheet7','JuvIPC':'Sheet8','JuvSLL':'Sheet9','Missing':'Sheet10'}

# ---------------------------------------------------------------------------
# 1. Canonical schema per category: list of (sub_indicator_id, label, [match substrings])
#    A sub_indicator's value = sum of every raw column whose flattened header
#    contains ANY of its match substrings (case-insensitive). This is robust
#    to the column-position drift we found between years.
# ---------------------------------------------------------------------------

# Each sub-indicator: (id, label, contains_patterns, exact_patterns)
#   contains_patterns -> substring match, used mainly for 2017-2019's verbose headers
#   exact_patterns    -> exact (whole, trimmed, lowercased) match, used for 2014-2016's
#                        short plain headers where substring matching would over-collect
#                        (e.g. "rape" is a substring of "attempt to commit rape" too)
CATEGORY_SCHEMA = {
    'Women': [
        ('rape', 'Rape', ['rape (sec'], ['rape']),
        ('attempt_rape', 'Attempt to commit Rape', ['attempt to commit rape'], ['attempt to commit rape']),
        ('kidnapping_abduction_total', 'Kidnapping & Abduction (Total)', ['kidnapping & abduction of women (total)', 'kidnapping and abduction of women (total)'], ['kidnapping & abduction_total']),
        ('dowry_deaths', 'Dowry Deaths', ['dowry death'], ['dowry deaths']),
        ('assault_outrage_modesty_total', 'Assault on Women (outrage modesty)', ['assault on women with intent to outrage her modesty (col'], ['assault on women with intent to outrage her modesty_total']),
        ('insult_modesty_total', 'Insult to Modesty of Women', ['insult to the modesty of women (col'], ['insult to the modesty of women_total']),
        ('cruelty_husband', 'Cruelty by Husband or Relatives', ['cruelty by husband'], ['cruelty by husband or his relatives']),
        ('importation_girls', 'Importation of Girls', ['importation of girls'], ['importation of girls from foreign country']),
        ('abetment_suicide', 'Abetment of Suicide of Women', ['abetment to suicide of women'], ['abetment of suicides of women']),
        ('dowry_prohibition_act', 'Dowry Prohibition Act', ['dowry prohibition act'], ['dowry prohibition act, 1961']),
        ('indecent_representation_act', 'Indecent Representation of Women Act', ['indecent representation of women'], ['indecent representation of women (p) act, 1986']),
        ('domestic_violence_act', 'Protection from Domestic Violence Act', ['protection of women from domestic violence'], ['protection of women from domestic violence act, 2005']),
        ('immoral_traffic_act', 'Immoral Traffic Prevention Act', ['immoral traffic'], ['immoral traffic prevention act']),
        ('acid_attack', 'Acid Attack', ['acid attack (sec'], ['acid attack']),
        ('attempt_acid_attack', 'Attempt to Acid Attack', ['attempt to acid attack'], ['attempt to acid attack']),
        ('miscarriage_total', 'Miscarriage (with/without consent)', ['miscarriage (sec'], ['deaths caused with intent to cause miscarriage', 'causing miscarriage without consent of women']),
        ('human_trafficking', 'Human Trafficking', ['human trafficking'], ['humantrafficking']),
        ('unnatural_offences', 'Unnatural Offences', ['unnatural offence'], ['unnatural offences']),
        ('cyber_crimes_women', 'Cyber Crimes against Women', ['cyber crimes/information technology act (women'], ['cyber crimes against women (section 67a it act)']),
        ('total_crimes_women', 'Total Crimes against Women', ['total crime against women'], ['total crimes against women']),
    ],
    'Children': [
        ('murder', 'Murder', ['murder (sec.302', 'murder (sec 302'], ['murder']),
        ('infanticide', 'Infanticide', ['infanticide'], ['infanticide']),
        ('rape', 'Rape', ['rape (sec. 376', 'rape (sec.376'], ['rape']),
        ('assault_outrage_modesty', 'Assault on Women (outrage modesty)', ['assault on women with intent to outrage her modesty'], ['assault on women with intent to outrage her modesty']),
        ('insult_modesty', 'Insult to Modesty of Women', ['insult to the modesty of women'], ['insult to the modesty of women']),
        ('kidnapping_abduction_total', 'Kidnapping & Abduction (Total)', ['kidnapping and abduction of children (sec'], ['kidnapping & abduction_total']),
        ('foeticide', 'Foeticide', ['foeticide'], ['foeticide']),
        ('abetment_suicide', 'Abetment of Suicide of Child', ['abetment of suicide of child', 'abetment to suicide of child'], ['abetment of suicide of child']),
        ('exposure_abandonment', 'Exposure and Abandonment', ['exposure and abandonment'], ['exposure and abandonment']),
        ('procuration_minor_girls', 'Procuration of Minor Girls', ['procuration of minor girls'], ['procuration of minor girls']),
        ('importation_girls', 'Importation of Girls', ['importation of girls'], ['importation of girls from foreign country']),
        ('buying_minors', 'Buying of Minors for Prostitution', ['buying of minors for prostitution (total', 'buying of minors for prostitution'], ['buying of minors for prostitution']),
        ('selling_minors', 'Selling of Minors for Prostitution', ['selling of minors for prostitution (total', 'selling of minors for prostitution'], ['selling of minors for prostitution']),
        ('child_marriage_act', 'Prohibition of Child Marriage Act', ['child marriage act'], ['prohibition of child marriage act, 2006']),
        ('organ_transplant_act', 'Transplantation of Human Organs Act', ['transplantation of human organs'], ['transplantation of human organs act, 1994']),
        ('child_labour_act', 'Child Labour Act', ['child labour'], ['child labour (prohibition & regulation) act, 1986']),
        ('immoral_traffic_act', 'Immoral Traffic Prevention Act', ['immoral traffic'], ['immoral traffic (prevention) act, 1956']),
        ('juvenile_justice_act', 'Juvenile Justice Act', ['juvenile justice', 'juveniles justice'], ['juveniles justice (care and protection of children) act, 2000']),
        ('pocso_act', 'POCSO Act (Total)', ['protection of children from sexual offences act (total'], ['protection of children from sexual offences act 2012']),
        ('attempt_murder', 'Attempt to Commit Murder', ['attempt to commit murder'], ['attempt to commit murder']),
        ('unnatural_offences', 'Unnatural Offences', ['r/w section 377 ipc'], ['unnatural offences']),
        ('human_trafficking', 'Human Trafficking', ['human trafficking'], ['human trafficking']),
        ('other_crimes', 'Other Crimes against Children', ['other ipc crimes', 'other sll crimes'], ['other crimes committed against children']),
        ('total_crimes_children', 'Total Crimes against Children', ['total crimes against children'], ['total crimes against children']),
    ],
    'SC': [
        ('pcr_act', 'Protection of Civil Rights Act', ['protection of civil rights act'], ['protection of civil rights act, 1955']),
        ('murder', 'Murder', ['murder (sec. 302', 'murder (sec.302'], ['poa_murder']),
        ('attempt_murder', 'Attempt to commit Murder', ['attempt to commit murder'], ['poa_attempt to commit murder']),
        ('rape', 'Rape', ['rape (sec. 376', 'rape (sec.376'], ['poa_rape']),
        ('attempt_rape', 'Attempt to commit Rape', ['attempt to commit rape'], ['poa_attempt to commit rape']),
        ('assault_outrage_modesty', 'Assault on Women (outrage modesty)', ['assault on women with intent to outrage her modesty'], ['poa_assault on women with intent to outrage her modesty']),
        ('insult_modesty', 'Insult to Modesty of Women', ['insult to the modesty of women'], ['poa_insult to the modesty of women']),
        ('kidnapping_abduction_total', 'Kidnapping & Abduction (Total)', ['kidnapping and abduction (sec'], ['poa_kidnapping & abduction_grandtotal']),
        ('dacoity', 'Dacoity', ['dacoity (sec'], ['poa_dacoity']),
        ('robbery', 'Robbery', ['robbery (sec'], ['poa_robbery']),
        ('arson', 'Arson', ['arson (sec'], ['poa_arson']),
        ('grievous_hurt', 'Grievous Hurt (Total)', ['grievous hurt (sec'], ['poa_grievous hurt']),
        ('riots', 'Riots/Rioting', ['rioting (sec', 'riots (sec'], ['poa_riots']),
        ('other_ipc', 'Other IPC crimes', ['other ipc crimes'], ['poa_other ipc crimes']),
        ('poa_act_only', 'SC/ST (POA) Act only', ['prevention of atrocities) act only'], ['poa_sc / st (prevention of atrocities) act only']),
        ('total_poa_act', 'Total of SC/ST (POA) Act', ['total of sc / st (prevention of atrocities) act', 'total of sc/st (prevention of atrocities) act'], ['total of sc/st (prevention of atrocities) act ,1989']),
        ('total_crimes_sc', 'Total Crimes against SCs', ['total crime/atrocities against scheduled castes'], ['total crimes against scs']),
    ],
    'ST': [
        ('pcr_act', 'Protection of Civil Rights Act', ['protection of civil rights act'], ['protection of civil rights act, 1955']),
        ('murder', 'Murder', ['murder (sec. 302', 'murder (sec.302'], ['poa_murder']),
        ('attempt_murder', 'Attempt to commit Murder', ['attempt to commit murder'], ['poa_attempt to commit murder']),
        ('rape', 'Rape', ['rape (sec. 376', 'rape (sec.376'], ['poa_rape']),
        ('attempt_rape', 'Attempt to commit Rape', ['attempt to commit rape'], ['poa_attempt to commit rape']),
        ('assault_outrage_modesty', 'Assault on Women (outrage modesty)', ['assault on women with intent to outrage her modesty'], ['poa_assault on women with intent to outrage her modesty']),
        ('insult_modesty', 'Insult to Modesty of Women', ['insult to the modesty of women'], ['poa_insult to the modesty of women']),
        ('kidnapping_abduction_total', 'Kidnapping & Abduction (Total)', ['kidnapping and abduction (sec'], ['poa_kidnapping & abduction_grandtotal']),
        ('dacoity', 'Dacoity', ['dacoity (sec'], ['poa_dacoity']),
        ('robbery', 'Robbery', ['robbery (sec'], ['poa_robbery']),
        ('arson', 'Arson', ['arson (sec'], ['poa_arson']),
        ('grievous_hurt', 'Grievous Hurt (Total)', ['grievous hurt (sec'], ['poa_grievous hurt']),
        ('riots', 'Riots/Rioting', ['rioting (sec', 'riots (sec'], ['poa_riots']),
        ('other_ipc', 'Other IPC crimes', ['other ipc crimes'], ['poa_other ipc crimes']),
        ('poa_act_only', 'SC/ST (POA) Act only', ['prevention of atrocities) act only'], ['poa_sc / st (prevention of atrocities) act only']),
        ('total_poa_act', 'Total of SC/ST (POA) Act', ['total of sc / st (prevention of atrocities) act', 'total of sc/st (prevention of atrocities) act'], ['total of sc/st (prevention of atrocities) act ,1989']),
        ('total_crimes_st', 'Total Crimes against STs', ['total crime/atrocities against scheduled tribes'], ['total crimes against sts']),
    ],
    'IPC': [
        ('murder', 'Murder', ['murder (sec.302'], ['murder']),
        ('attempt_murder', 'Attempt to commit Murder', ['attempt to commit murder (sec.307'], ['attempt to commit murder']),
        ('culpable_homicide', 'Culpable Homicide not amounting to Murder', ['culpable homicide not amounting to murder'], ['culpable homicide not amounting to murder']),
        ('attempt_culpable_homicide', 'Attempt to commit Culpable Homicide', ['attempt to commit culpable homicide'], ['attempt to commit culpable homicide']),
        ('rape', 'Rape', ['rape (sec.376 ipc)'], ['rape']),
        ('attempt_rape', 'Attempt to commit Rape', ['attempt to commit rape'], ['attempt to commit rape']),
        ('kidnapping_abduction_total', 'Kidnapping & Abduction (Total)', ['kidnapping and abduction (total)'], ['kidnapping & abduction_total']),
        ('dacoity', 'Dacoity', ['dacoity  (total', 'dacoity (total'], ['dacoity']),
        ('dacoity_prep', 'Preparation/Assembly for Dacoity', ['making preparation and assembly for committing dacoity'], ['making preparation and assembly for committing dacoity']),
        ('robbery', 'Robbery', ['robbery  (sec.392', 'robbery (sec.392'], ['robbery']),
        ('trespass_burglary', 'Criminal Trespass/Burglary', ['criminal trespass (sec', 'burglary (total'], ['criminal trespass/burglary']),
        ('theft', 'Theft', ['theft (total'], ['theft']),
        ('unlawful_assembly', 'Unlawful Assembly', ['unlawful assembly (sec'], ['unlawful assembly']),
        ('riots', 'Riots', ['rioting (total'], ['riots']),
        ('breach_of_trust', 'Criminal Breach of Trust', ['criminal breach of trust'], ['criminal breach of trust']),
        ('cheating', 'Cheating', ['cheating (sec.420'], ['cheating']),
        ('forgery', 'Forgery', [], ['forgery']),  # forced-exact even in nested years, see FORCE_EXACT_IDS
        ('counterfeiting', 'Counterfeiting', ['counterfeiting (total'], ['counterfeiting']),
        ('arson', 'Arson', ['arson (sec.435'], ['arson']),
        ('grievous_hurt', 'Grievous Hurt', ['grievous hurt'], ['grievous hurt']),
        ('dowry_deaths', 'Dowry Deaths', ['dowry deaths (sec'], ['dowry deaths']),
        ('assault_outrage_modesty', 'Assault on Women (outrage modesty)', ['outrage her modesty (sec.354 ipc) (total'], ['assault on women with intent to outrage her modesty']),
        ('insult_modesty', 'Insult to Modesty of Women', ['insult to the modesty of women (sec'], ['insult to the modesty of women']),
        ('cruelty_husband', 'Cruelty by Husband or Relatives', ['cruelty by husband or his relatives (sec'], ['cruelty by husband or his relatives']),
        ('importation_girls', 'Importation of Girls', ['importation of girls from foreign country'], ['importation of girls from foreign country']),
        ('death_by_negligence', 'Causing Death by Negligence', ['causing death by negligence (sec.304-a'], ['causing death by negligence']),
        ('offences_against_state', 'Offences against State', ['offences against state (total'], ['offences against state']),
        ('offences_promoting_enmity', 'Offences promoting enmity between groups', ['offences promoting enmity between different groups (total'], ['offences promoting enmity between different groups']),
        ('extortion', 'Extortion', ['extortion & blackmailing'], ['extortion']),
        ('disclosure_victim_identity', 'Disclosure of Identity of Victims', [], ['disclosure of identity of victims']),  # genuine gap post-2016
        ('rash_driving', 'Incidence of Rash Driving', ['rash driving on public way (total'], ['incidence of rash driving']),
        ('human_trafficking', 'HumanTrafficking', ['human trafficking (u/s 370'], ['humantrafficking']),
        ('unnatural_offence', 'Unnatural Offence', ['unnatural offences (sec'], ['unnatural offence']),
        ('other_ipc', 'Other IPC crimes', ['other ipc crimes'], ['other ipc crimes']),
        ('total_ipc', 'Total Cognizable IPC crimes', ['total cognizable ipc crimes'], ['total cognizable ipc crimes']),
    ],
    'SLL': [
        ('arms_act', 'Arms Act', ['the arms act, 1959 (total'], ['arms act']),
        ('ndps_act', 'NDPS Act', ['the ndps act, 1985 (total'], ['ndps  act', 'ndps act']),
        ('gambling_act', 'Gambling Act', ['the gambling act'], ['gambling act']),
        ('excise_act', 'Excise Act', ['the excise act'], ['excise act']),
        ('prohibition_act', 'Prohibition Act', ['prohibition act (state)'], ['prohibition act']),
        ('explosives_act', 'Explosives & Explosive Substances Act', ['the explosives act', 'the explosive substances act'], ['explosives & explosive substances act']),
        ('immoral_traffic_act', 'Immoral Traffic Prevention Act', ['the immoral traffic prevention act'], ['immoral traffic (prevention) act']),
        ('railways_act', 'Indian Railways Act', ['the indian railways act'], ['indian railways act']),
        ('foreigners_act', 'Registration of Foreigners Act + Foreigners Act (merged)', ['the registration of foreigners act'], ['registration of foreigners act', 'foreigners act']),
        ('pcr_act', 'Protection of Civil Rights Act', ['protection of civil rights act, 1955 against'], ['protection of civil rights act']),
        ('passport_act', 'Passport Act', ['the passport act'], ['passport act']),
        ('essential_commodities_act', 'Essential Commodities Act', ['the essential commodities act'], ['essential commodities act']),
        ('antiquities_act', 'Antiquities & Art Treasures Act', ['the antiques and art treasures act'], ['antiquities & art treasures act']),
        ('dowry_prohibition_act', 'Dowry Prohibition Act', ['the dowry prohibition act'], ['dowry prohibition act']),
        ('indecent_representation_act', 'Indecent Representation of Women Act', ['the indecent representation of women prohibition act'], ['indecent representation of women (prohibition) act']),
        ('copyright_act', 'Copyright Act', ['the copy right act'], ['copyright act']),
        ('sati_prevention_act', 'Commission of Sati Prevention Act', [], ['commission of sati prevention act']),  # genuine gap
        ('sc_st_poa_act', 'SC/ST (POA) Act', ['the sc/st prevention of atrocities act'], ['sc/st (poa) act']),
        ('forest_act', 'Forest Act', ['the forest act, 1927'], ['forest act']),
        ('child_marriage_act', 'Prohibition of Child Marriage Act', ['the prohibition of child marriage act'], ['prohibition of child marriage act']),
        ('domestic_violence_act', 'Protection from Domestic Violence Act', ['the protection of women from domestic violence act'], ['protection of women from domestic violence act']),
        ('it_act', 'Information Technology Act', ['the information technology act'], ['information technology act']),
        ('official_secrets_act', 'Official Secrets Act', ['the official secrets act'], ['official secrets act']),
        ('electricity_act', 'Electricity Act', ['the electricity act'], ['electricity act']),
        ('wildlife_act', 'Wildlife Protection Act', ['the wildlife protection act'], ['wildlife protection act']),
        ('bonded_labour_act', 'Bonded Labour System Act', ['the bonded labour system abolition act'], ['bonded labour system (abolition) act']),
        ('air_water_pollution_act', 'Air + Water Pollution Act (merged)', ['the air & the water prevention'], ['air (prevention & control of pollution) act', 'water (prevention & control of pollution) act']),
        ('national_security_act', 'National Security Act', [], ['national security act']),  # genuine gap
        ('unlawful_activities_act', 'Unlawful Activities Act', ['the unlawful activities p act'], ['unlawful activities (prevention) act']),
        ('young_persons_act', 'Young Persons (Harmful Publications) Act', [], ['young persons (harmful publications) act']),  # genuine gap
        ('railway_property_act', 'Railway Property Act', ['the railway property unlawful possession act'], ['railway property (unlawful possession) act']),
        ('public_property_act', 'Prevention of Damage to Public Property Act', ['the prevention of damage to public property act'], ['prevention of damage to public property act']),
        ('organ_transplant_act', 'Transplantation of Human Organs Act', ['the transplantation of human organs act'], ['transplantation of human organs act']),
        ('trade_marks_act', 'Trade Marks Act', ['the trade marks act'], ['trade marks act']),
        ('national_honour_act', 'Prevention of Insults to National Honour Act', ['the prevention of insults to national honour act'], ['prevention of insults to national honour act']),
        ('state_emblem_act', 'State Emblem Act', [], ['state emblem  (prohibition of improper use) act']),  # genuine gap
        ('lotteries_act', 'Lotteries Act', ['the lotteries regulation act'], ['lotteries (regulation) act']),
        ('citizenship_act', 'Citizenship Act', ['the citizenship act'], ['citizenship act']),
        ('worship_act', 'Place of Worship Act', [], ['place of worship (special provisions) act']),  # genuine gap
        ('religious_institution_act', 'Religious Institution Act', [], ['religious institution (prevention of misuse) act']),  # genuine gap
        ('representation_people_act', 'Representation of the People Act', ['the representation of the people act'], ['representation of the people act']),
        ('emigration_act', 'Emigration Act', ['the emigration act'], ['emigration act']),
        ('juvenile_justice_act', 'Juvenile Justice Act', ['the juvenile justice care and protection'], ['juvenile justice (care and protection of children) act']),
        ('infant_substitutes_act', 'Infant Substitutes Regulation Act', [], ['infant substitutes regulation act']),  # genuine gap
        ('anti_hijacking_act', 'Anti-Hijacking Act', [], ['anti-hijacking act']),  # genuine gap
        ('atomic_energy_act', 'Atomic Energy Act', [], ['atomic energy act']),  # genuine gap
        ('wmd_act', 'Weapons of Mass Destruction Act', [], ['weapons of mass destruction (proh of unlawful activities) act']),  # genuine gap
        ('civil_aviation_act', 'Safety of Civil Aviation Act', [], ['supprn of unlawful acts against safety of civil aviation act']),  # genuine gap
        ('maritime_navigation_act', 'Safety of Maritime Navigation Act', [], ['safety of maritime navigation act']),  # genuine gap
        ('manual_scavengers_act', 'Manual Scavengers Act', [], ['manual scavengers and construction of dry latrines (p) act']),  # genuine gap
        ('prenatal_diagnostic_act', 'Pre-Natal Diagnostic Techniques Act', ['the pre-natal diagnostic techniques reg and prev of misuse act'], ['pre-natal diagnostic techniques (reg and prev of misuse) act']),
        ('maritime_zone_act', 'Maritime Zone Act', [], ['the maritime zone of india (reg of fishing by for vessel) act']),  # genuine gap
        ('other_sll', 'Other SLL crimes', ['other sll crimes'], ['other sll crimes']),
        ('total_sll', 'Total Cognizable SLL crimes', ['total cognizable sll crimes'], ['total cognizable sll crimes']),
    ],
    'Cyber': [
        ('total_it_act', 'Total Offences under IT Act', ['total offences under i.t. act'], []),
        ('total_ipc_rw_it', 'Total Offences under IPC r/w IT Act', ['total offences under ipc r/w it act'], []),
        ('total_sll_rw_it', 'Total Offences under SLL r/w IT Act', ['total offences under sll r/w it act'], []),
        ('total_cyber', 'Total Cyber Crimes', ['total cyber crimes'], []),
    ],
    'Missing': [
        ('total_missing', 'Total Missing Persons', ['grand total'], []),
    ],
}

# sub-indicator ids that need EXACT matching even for nested-format years
# (their text is a short substring of another column's longer header, so
# "contains" matching would wrongly grab both)
FORCE_EXACT_IDS = {'forgery'}

# ---------------------------------------------------------------------------
# 2. Header extraction
# ---------------------------------------------------------------------------

def flat_headers_nested(path, sheet):
    """Return list of (col_index, flattened_header_text) for 2017-2019 style files."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=4)
    out = []
    for c in range(raw.shape[1]):
        parts = [str(raw.iloc[r, c]).strip() for r in range(1, 4) if str(raw.iloc[r, c]) not in ('nan', 'None')]
        out.append((c, " | ".join(parts)))
    return out

def flat_headers_simple(path, sheet):
    """Return list of (col_index, header_text) for 2014-2016 style files."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=2)
    out = []
    for c in range(raw.shape[1]):
        v = str(raw.iloc[1, c]).strip()
        if v not in ('nan', 'None'):
            out.append((c, v))
    return out

def match_columns_contains(headers, patterns):
    """Return list of col indices whose header text CONTAINS any pattern (case-insensitive)."""
    matched = []
    for col, text in headers:
        low = text.lower()
        if any(p in low for p in patterns):
            matched.append(col)
    return matched

def match_columns_exact(headers, patterns):
    """Return list of col indices whose header text EXACTLY equals any pattern (case-insensitive, trimmed)."""
    matched = []
    for col, text in headers:
        low = text.strip().lower()
        if low in patterns:
            matched.append(col)
    return matched

# ---------------------------------------------------------------------------
# 3. Crosswalk
# ---------------------------------------------------------------------------

def load_crosswalk():
    cw = pd.read_csv('ncrb_district_crosswalk.csv')
    lookup = {}
    for _, r in cw.iterrows():
        key = (str(r['ncrb_state_raw']).strip().upper(), str(r['ncrb_district_raw']).strip().upper())
        if pd.notna(r['district_code_pc11']) and pd.notna(r['state_code']):
            lookup[key] = (int(r['state_code']), int(r['district_code_pc11']))
    return lookup

# ---------------------------------------------------------------------------
# 4. Main extraction per year/category
# ---------------------------------------------------------------------------

def process_year_category(year, category, crosswalk):
    fmt = 'simple' if year in SIMPLE_FORMAT_YEARS else 'nested'
    sheet_map = SHEET_MAP_SIMPLE if fmt == 'simple' else SHEET_MAP_NESTED
    if category not in sheet_map:
        return [], [], 0  # category doesn't exist for this year (e.g. Cyber pre-2017)

    path = f'All{year}.xlsx'
    sheet = sheet_map[category]

    if fmt == 'nested':
        headers = flat_headers_nested(path, sheet)
        df = pd.read_excel(path, sheet_name=sheet, header=None, skiprows=4)
    else:
        headers = flat_headers_simple(path, sheet)
        df = pd.read_excel(path, sheet_name=sheet, header=None, skiprows=2)

    schema = CATEGORY_SCHEMA.get(category, [])
    if category == 'JuvIPC':
        schema = CATEGORY_SCHEMA['IPC']
    elif category == 'JuvSLL':
        schema = CATEGORY_SCHEMA['SLL']

    unmatched_schema = []
    match_fn = match_columns_exact if fmt == 'simple' else match_columns_contains

    # Both formats use the same layout: no per-row state column. Instead, a
    # standalone marker row ("State: X" or "State : X") precedes each state's
    # block of districts, with the district name in column 1. Footnote/junk
    # rows (e.g. "* Due to non-receipt of data...") appear too and are skipped
    # since they won't match any crosswalk key.
    row_geo = []  # (row_index_in_df, state_name, district_name)
    current_state = None
    for idx, row in df.iterrows():
        col0 = str(row[0]).strip() if pd.notna(row[0]) else ''
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ''
        if col0.upper().replace(' ', '').startswith('STATE:') or col0.upper().startswith('STATE :'):
            current_state = re.split(r'STATE\s*:', col0.upper(), maxsplit=1)[-1].strip()
            continue
        if col1 == '' or current_state is None:
            continue
        if 'TOTAL' in col1.upper():
            continue
        row_geo.append((idx, current_state, col1.upper()))

    rows_out = []
    unresolved_geo = 0
    for sub_id, label, contains_patterns, exact_patterns in schema:
        if fmt == 'simple' or sub_id in FORCE_EXACT_IDS:
            cols = match_columns_exact(headers, exact_patterns)
        else:
            cols = match_fn(headers, contains_patterns)
        if not cols:
            unmatched_schema.append((sub_id, label))
            continue
        for idx, raw_state, raw_dist in row_geo:
            key = (raw_state, raw_dist)
            if key not in crosswalk:
                unresolved_geo += 1
                continue
            sc, dc = crosswalk[key]
            row = df.loc[idx]
            try:
                val = sum(float(row[c]) for c in cols if pd.notna(row[c]) and str(row[c]).strip() not in ('', '-'))
            except (ValueError, TypeError):
                continue
            rows_out.append({
                'state_code': sc, 'district_code_pc11': dc,
                'sub_indicator': sub_id, 'year': year, 'value': val
            })

    return rows_out, unmatched_schema, unresolved_geo

# ---------------------------------------------------------------------------
# 5. Run
# ---------------------------------------------------------------------------

def main():
    crosswalk = load_crosswalk()
    print(f"Crosswalk loaded: {len(crosswalk)} district-name entries")

    os.makedirs('../data_out', exist_ok=True)

    for category in ['IPC', 'SLL', 'Women', 'Children', 'SC', 'ST', 'Cyber', 'JuvIPC', 'JuvSLL', 'Missing']:
        schema = CATEGORY_SCHEMA['IPC'] if category == 'JuvIPC' else \
                 CATEGORY_SCHEMA['SLL'] if category == 'JuvSLL' else \
                 CATEGORY_SCHEMA[category]
        all_rows = []
        print(f"\n=== {category} ===")
        for year in YEARS:
            rows, unmatched, unresolved_geo = process_year_category(year, category, crosswalk)
            all_rows.extend(rows)
            status = f"{len(rows)} rows"
            if unmatched:
                status += f" | UNMATCHED SCHEMA: {[u[0] for u in unmatched]}"
            if unresolved_geo:
                status += f" | unresolved district keys: {unresolved_geo}"
            print(f"  {year}: {status}")

        sub_indicators = [{'id': sid, 'label': lbl} for sid, lbl, _, _ in schema]
        out = {'category': category, 'sub_indicators': sub_indicators, 'data': all_rows}
        with open(f'../data_out/{category.lower()}.json', 'w') as f:
            json.dump(out, f)
        print(f"  -> wrote {len(all_rows)} total rows to {category.lower()}.json")

if __name__ == '__main__':
    main()
