

hi!
some quick notes for ease


how i did the extraction stuff:


a given dataset name looks like:
20251120_abc_abc_f03_aaaaa_c-111_Pilot1_hq_no_bbera
The code does:

Date: name.split('_')[0] → takes the first chunk before any underscore → 20251120
Pilot: re.search(r'(Pilot\d+)', name) → scans for "Pilot" followed by digits anywhere in the string → Pilot1
Config: re.findall(r'(c-\d+)', name)[-1] → finds ALL matches of "c-" followed by digits, takes the last one (because the episode ID sometimes contains a c-XXX too) → c-111
Quality: '_hq' in name → simple substring check → hq
No BB: 'no_bb' in name.lower() → substring check → yes
Speed: 'speed' in name.lower() → substring check → no


regex use cases:

1. for analysis/charts:
lot_match = re.search(r'(Pilot\d+)', name)
config_matches = re.findall(r'(c-\d+)', name)


2: excel wkbk
pilot_match = re.search(r'(Pilot\d+)', name)
config_matches = re.findall(r'(c-\d+)', name)


important:

both scripts import re at the top and use the same parse function. The two regex patterns are:

r'(Pilot\d+)' — matches "Pilot" followed by one or more digits → extracts Pilot1, Pilot2, etc.
r'(c-\d+)' — matches "c-" followed by one or more digits → extracts c-100, c-111, c-198, etc.

The other fields (quality, no_bb, speed) use simple string matching ('_hq' in name), not regex
