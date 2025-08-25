# -*- coding: utf-8 -*-
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import re
import calendar
import unicodedata
import logging

# Ρύθμιση logging
logging.basicConfig(level=logging.DEBUG)

# Συνάρτηση για καθαρισμό κειμένου (πεζά, χωρίς τόνους/διαλυτικά)
def normalize_text(text):
    if pd.isna(text):
        return ""
    text = ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn')
    return text.lower().strip()

# Φόρτωση mapping_ενεργειας.xlsx
try:
    mapping_df = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='mapping', engine='openpyxl')
    eidos_rules = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='Eidos_Rules', engine='openpyxl')
    config_df = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='Config', engine='openpyxl')
    file_mapping_df = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='FileMapping', engine='openpyxl')
    partner_rules_df = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='PartnerRules', engine='openpyxl')
    date_formats_df = pd.read_excel('mapping_ενεργειας.xlsx', sheet_name='DateFormats', engine='openpyxl')
except FileNotFoundError:
    print("❌ Το αρχείο mapping_ενεργειας.xlsx δεν βρέθηκε στον φάκελο.")
    exit()
except ValueError as e:
    print(f"❌ Σφάλμα στη δομή του mapping_ενεργειας.xlsx: {str(e)}")
    exit()
except Exception as e:
    print(f"❌ Σφάλμα κατά τη φόρτωση του mapping_ενεργειας.xlsx: {str(e)}")
    exit()

# Έλεγχος ύπαρξης στήλης Setting στο Config
if 'Setting' not in config_df.columns or 'Value' not in config_df.columns:
    print("❌ Το sheet Config δεν περιέχει τις στήλες 'Setting' ή 'Value'.")
    print(f"Στήλες Config: {list(config_df.columns)}")
    exit()

# Έλεγχος ύπαρξης στήλης FilePattern στο FileMapping
if 'Provider' not in file_mapping_df.columns or 'FilePattern' not in file_mapping_df.columns:
    print("❌ Το sheet FileMapping δεν περιέχει τις στήλες 'Provider' ή 'FilePattern'.")
    print(f"Στήλες FileMapping: {list(file_mapping_df.columns)}")
    exit()

# Δημιουργία ρυθμίσεων από Config
config = {}
for setting in config_df['Setting'].unique():
    values = config_df[config_df['Setting'] == setting]['Value'].dropna().tolist()
    if setting == 'IgnoreKeywords':
        config[setting] = [item for sublist in [v.split(',') for v in values] for item in sublist]
    else:
        config[setting] = values if len(values) > 1 else values[0] if values else ''
ignore_keywords = config.get('IgnoreKeywords', [])
provider_keywords = config.get('Providers', [])
output_file = config.get('OutputFile', 'πίνακας αμοιβών %s.xlsx')
date_columns = config.get('DateColumns', [])
valid_sheets = {}
for pair in config.get('ValidSheetsPerProvider', []):
    if ':' in pair:
        provider, sheets = pair.split(':')
        valid_sheets[provider] = sheets.split(',')
fpa_rate = float(config.get('FPA_Rate', 0.24))

# Δημιουργία mapping_dict
new_column_names = mapping_df.iloc[:, 0].dropna().tolist()
mapping_dict = {}
for col in mapping_df.columns[1:]:
    provider_map = {}
    for i, val in enumerate(mapping_df[col]):
        if pd.notna(val):
            provider_map[str(val).strip()] = mapping_df.iloc[i, 0]
    mapping_dict[col.strip()] = provider_map

# Συνάρτηση για επικύρωση ημερομηνίας
def validate_date(date_str):
    pattern = r'^\d{4}-\d{2}$'
    if not re.match(pattern, date_str):
        return False
    try:
        year, month = map(int, date_str.split('-'))
        if not (1 <= month <= 12):
            return False
        datetime(year, month, 1)
        return True
    except ValueError:
        return False

# Συνάρτηση για μετατροπή YYYY-MM σε datetime dd/mm/yyyy
def get_last_day_of_month(date_str):
    year, month = map(int, date_str.split('-'))
    _, last_day = calendar.monthrange(year, month)
    return datetime(year, month, last_day).strftime('%d/%m/%Y')

# Συνάρτηση για καθαρισμό Partner
def clean_partner(value, errors, partner_rules_df):
    if pd.isna(value) or not isinstance(value, str):
        return value
    for _, rule in partner_rules_df.iterrows():
        pattern, replacement = rule['Pattern'], rule['Replacement']
        if re.search(pattern, value, re.IGNORECASE):
            return re.sub(pattern, replacement, value, flags=re.IGNORECASE)
    return value

# Συνάρτηση για καθορισμό table_id
def get_table_id(columns, provider_key, sheet_name, valid_sheets, table_title):
    if "Zenith_R" in provider_key and sheet_name in valid_sheets.get("Zenith_R", []):
        return provider_key
    if provider_key in ["NRG", "NLS_ΗΡΩΝ"]:
        return provider_key
    columns_lower = [normalize_text(col) for col in columns if pd.notna(col)]
    if provider_key == "PROTERGIA":
        if any(re.search(r'προιον', col, re.IGNORECASE) for col in columns_lower):
            return "PROTERGIA1"
        return "Any"
    elif provider_key == "ΖΕΝΊΘ":
        if any(re.search(r'αμοιβη 1ης συμβασης', col, re.IGNORECASE) for col in columns_lower):
            return "ZENITH1"
        return "Any"
    elif provider_key == "ELPEDISON":
        if any(re.search(r'supply category', col, re.IGNORECASE) for col in columns_lower):
            return "ELPEDISON1"
        return "Any"
    elif provider_key in ["VOLTON", "ΗΡΩΝ"]:
        return "Any"
    return "Any"

# Συνάρτηση για έλεγχο αν μια γραμμή είναι τίτλος πίνακα
def is_header_row(row, provider, sheet_name=None):
    row_values = [normalize_text(cell) for cell in row if pd.notna(cell)]
    for key in mapping_dict:
        if provider.lower() in key.lower() and (provider != "Zenith_R" or key == f"Zenith_R{sheet_name}"):
            expected_headers = [normalize_text(h) for h in mapping_dict[key].keys()]
            matching_count = sum(1 for val in row_values if val in expected_headers)
            if len(row_values) >= 2 and matching_count >= len(row_values) - 1:
                return True, key
    return False, None

# Συνάρτηση για καθορισμό Είδους
def set_eidos(row, provider, provider_key, table_id, errors, columns, eidos_rules, sheet_name=None, table_title=None):
    try:
        if all(pd.isna(cell) or str(cell).strip() == "" for cell in row):
            return ""

        provider_rules = eidos_rules[eidos_rules['Πάροχος'] == provider]
        if provider_rules.empty:
            errors.append(f"No Eidos rules found for provider {provider}")
            return ""

        provider_rules = provider_rules.sort_values(by='Προτεραιότητα')
        normalized_columns = [normalize_text(col) for col in columns if pd.notna(col)]

        for _, rule in provider_rules.iterrows():
            rule_table = rule['Πίνακας']
            if pd.isna(rule_table):
                rule_table = 'Any'
            condition = rule['Συνθήκη']
            eidos_value = rule['Τιμή Είδους']
            condition_type = rule.get('ConditionType', 'None')
            if pd.isna(condition_type):
                condition_type = 'None'

            if rule_table != 'Any' and table_id not in rule_table.split('|'):
                continue

            normalized_condition = normalize_text(condition)

            conditions_met = True

            if not isinstance(condition, str):
                condition = ''

            sub_conditions = condition.split(' and ')
            sub_types = condition_type.split('_') if '_' in condition_type else [condition_type] * len(sub_conditions)

            if len(sub_conditions) > 1:
                for sub_idx, sub_cond in enumerate(sub_conditions):
                    sub_cond_norm = normalize_text(sub_cond.strip())
                    sub_type = sub_types[sub_idx] if sub_idx < len(sub_types) else 'None'

                    match = True

                    if sub_type == 'sheet' and sub_cond.startswith('sheet: '):
                        sheet_num = sub_cond.split('sheet: ')[1].strip()
                        if sheet_num != sheet_name:
                            match = False
                    elif sub_type == 'contains':
                        if sub_cond_norm.endswith(' in title'):
                            key_phrase = normalize_text(sub_cond_norm.split(' in title')[0])
                            if not (table_title and key_phrase in normalize_text(table_title)):
                                match = False
                        else:
                            # Για άλλες contains (όχι in title), εφαρμόστε γενικό contains σε στήλες
                            try:
                                col, value = sub_cond_norm.split(' = ')
                                col_norm = normalize_text(col)
                                if col_norm in df.columns:
                                    row_value = normalize_text(row[col_norm])
                                    if value not in row_value:
                                        match = False
                                else:
                                    match = False
                            except ValueError:
                                match = False  # Αν η συνθήκη δεν έχει μορφή 'στήλη = τιμή'
                    elif sub_type == 'startswith' and sub_cond_norm == 'αμοιβη 1ης συμβασης in title and τυπος τιμολογιου starts with γ':
                        if table_id == 'ZENITH1' and any(re.search(r'τυπος τιμολογιου', col, re.IGNORECASE) for col in normalized_columns):
                            idx = [i for i, col in enumerate(normalized_columns) if re.search(r'τυπος τιμολογιου', col, re.IGNORECASE)][0]
                            type_value = normalize_text(row[columns[idx]])
                            if not (pd.notna(row[columns[idx]]) and isinstance(row[columns[idx]], str) and type_value.startswith('γ')):
                                match = False
                    elif sub_type == 'contains' and sub_cond_norm == 'αμοιβη 1ης συμβασης in title':
                        if table_id != 'ZENITH1':
                            match = False
                    elif sub_type == 'numeric' and sub_cond_norm.endswith(' is numeric'):
                        column_name = sub_cond_norm.split(' is numeric')[0]
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'καταναλωση|ενεργεια|kwh|mwh', normalize_text(col), re.IGNORECASE):
                                try:
                                    value = pd.to_numeric(row[columns[i]], errors='coerce')
                                    if pd.notna(value):
                                        found = True
                                        break
                                except:
                                    pass
                        if not found:
                            match = False
                    elif sub_type == 'startswith' and sub_cond_norm.endswith(' starts with γ'):
                        column_name = sub_cond_norm.split(' starts with γ')[0]
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(column_name, normalize_text(col), re.IGNORECASE):
                                type_value = normalize_text(row[columns[i]])
                                if pd.notna(row[columns[i]]) and isinstance(row[columns[i]], str) and type_value.startswith('γ'):
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'not_startswith' and sub_cond_norm.endswith(' not starts with γ'):
                        column_name = sub_cond_norm.split(' not starts with γ')[0]
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(column_name, normalize_text(col), re.IGNORECASE):
                                type_value = normalize_text(row[columns[i]])
                                if pd.isna(row[columns[i]]) or (isinstance(row[columns[i]], str) and not type_value.startswith('γ')):
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'exists' and sub_cond_norm.endswith(' exists'):
                        column_name = sub_cond_norm.split(' exists')[0]
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(column_name, normalize_text(col), re.IGNORECASE):
                                if pd.notna(row[columns[i]]) and str(row[columns[i]]).strip():
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'equals' and sub_cond_norm == 'προιον = electricity':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'προιον', col, re.IGNORECASE):
                                if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'electricity':
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'equals' and sub_cond_norm == 'προιον = gas':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'προιον', col, re.IGNORECASE):
                                if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'gas':
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'equals' and sub_cond_norm == 'κατασταση παροχης = terminated':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'κατασταση παροχης', col, re.IGNORECASE):
                                if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'terminated':
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'equals' and sub_cond_norm == 'αμοιβη = 5':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'αμοιβη|προμηθεια|commission', col, re.IGNORECASE):
                                try:
                                    if pd.notna(row[columns[i]]) and float(row[columns[i]]) == 5:
                                        found = True
                                        break
                                except (ValueError, TypeError):
                                    pass
                        if not found:
                            match = False
                    elif sub_type == 'equals' and sub_cond_norm == 'αμοιβη = 0.10':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'αμοιβη|προμηθεια|commission', col, re.IGNORECASE):
                                try:
                                    if pd.notna(row[columns[i]]) and float(row[columns[i]]) == 0.10:
                                        found = True
                                        break
                                except (ValueError, TypeError):
                                    pass
                        if not found:
                            match = False
                    elif sub_type == 'startswith' and sub_cond_norm == 'τυπος παροχης starts with γ':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'τυπος παροχης', col, re.IGNORECASE):
                                type_value = normalize_text(row[columns[i]])
                                if pd.notna(row[columns[i]]) and isinstance(row[columns[i]], str) and type_value.startswith('γ'):
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'not_startswith' and sub_cond_norm == 'τυπος παροχης not starts with γ':
                        found = False
                        for i, col in enumerate(normalized_columns):
                            if re.search(r'τυπος παροχης', col, re.IGNORECASE):
                                type_value = normalize_text(row[columns[i]])
                                if pd.isna(row[columns[i]]) or (isinstance(row[columns[i]], str) and not type_value.startswith('γ')):
                                    found = True
                                    break
                        if not found:
                            match = False
                    elif sub_type == 'None' or sub_cond == '':
                        pass  # Fallback handled below
                    else:
                        match = False

                    if not match:
                        conditions_met = False
                        break
            else:
                # Single condition handling
                match = True
                if condition_type == 'sheet' and condition.startswith('sheet: '):
                    sheet_num = condition.split('sheet: ')[1].strip()
                    if sheet_num != sheet_name:
                        match = False
                elif condition_type == 'contains' and normalized_condition == 'e-bill in title':
                    if not (table_title and 'e-bill' in normalize_text(table_title)):
                        match = False
                elif condition_type == 'contains' and normalized_condition == 'ebill in title':
                    if not (table_title and 'ebill' in normalize_text(table_title)):
                        match = False
                elif condition_type == 'startswith' and normalized_condition == 'αμοιβη 1ης συμβασης in title and τυπος τιμολογιου starts with γ':
                    if table_id == 'ZENITH1' and any(re.search(r'τυπος τιμολογιου', col, re.IGNORECASE) for col in normalized_columns):
                        idx = [i for i, col in enumerate(normalized_columns) if re.search(r'τυπος τιμολογιου', col, re.IGNORECASE)][0]
                        type_value = normalize_text(row[columns[idx]])
                        if not (pd.notna(row[columns[idx]]) and isinstance(row[columns[idx]], str) and type_value.startswith('γ')):
                            match = False
                elif condition_type == 'contains' and normalized_condition == 'αμοιβη 1ης συμβασης in title':
                    if table_id != 'ZENITH1':
                        match = False
                elif condition_type == 'numeric' and normalized_condition.endswith(' is numeric'):
                    column_name = normalized_condition.split(' is numeric')[0]
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'καταναλωση|ενεργεια|kwh|mwh', normalize_text(col), re.IGNORECASE):
                            try:
                                value = pd.to_numeric(row[columns[i]], errors='coerce')
                                if pd.notna(value):
                                    found = True
                                    break
                            except:
                                pass
                    if not found:
                        match = False
                elif condition_type == 'startswith' and normalized_condition.endswith(' starts with γ'):
                    column_name = normalized_condition.split(' starts with γ')[0]
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(column_name, normalize_text(col), re.IGNORECASE):
                            type_value = normalize_text(row[columns[i]])
                            if pd.notna(row[columns[i]]) and isinstance(row[columns[i]], str) and type_value.startswith('γ'):
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'not_startswith' and normalized_condition.endswith(' not starts with γ'):
                    column_name = normalized_condition.split(' not starts with γ')[0]
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(column_name, normalize_text(col), re.IGNORECASE):
                            type_value = normalize_text(row[columns[i]])
                            if pd.isna(row[columns[i]]) or (isinstance(row[columns[i]], str) and not type_value.startswith('γ')):
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'exists' and normalized_condition.endswith(' exists'):
                    column_name = normalized_condition.split(' exists')[0]
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(column_name, normalize_text(col), re.IGNORECASE):
                            if pd.notna(row[columns[i]]) and str(row[columns[i]]).strip():
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'equals' and normalized_condition == 'προιον = electricity':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'προιον', col, re.IGNORECASE):
                            if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'electricity':
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'equals' and normalized_condition == 'προιον = gas':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'προιον', col, re.IGNORECASE):
                            if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'gas':
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'equals' and normalized_condition == 'κατασταση παροχης = terminated':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'κατασταση παροχης', col, re.IGNORECASE):
                            if pd.notna(row[columns[i]]) and normalize_text(row[columns[i]]) == 'terminated':
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'equals' and normalized_condition == 'αμοιβη = 5':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'αμοιβη|προμηθεια|commission', col, re.IGNORECASE):
                            try:
                                if pd.notna(row[columns[i]]) and float(row[columns[i]]) == 5:
                                    found = True
                                    break
                            except (ValueError, TypeError):
                                pass
                    if not found:
                        match = False
                elif condition_type == 'equals' and normalized_condition == 'αμοιβη = 0.10':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'αμοιβη|προμηθεια|commission', col, re.IGNORECASE):
                            try:
                                if pd.notna(row[columns[i]]) and float(row[columns[i]]) == 0.10:
                                    found = True
                                    break
                            except (ValueError, TypeError):
                                pass
                    if not found:
                        match = False
                elif condition_type == 'startswith' and normalized_condition == 'τυπος παροχης starts with γ':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'τυπος παροχης', col, re.IGNORECASE):
                            type_value = normalize_text(row[columns[i]])
                            if pd.notna(row[columns[i]]) and isinstance(row[columns[i]], str) and type_value.startswith('γ'):
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'not_startswith' and normalized_condition == 'τυπος παροχης not starts with γ':
                    found = False
                    for i, col in enumerate(normalized_columns):
                        if re.search(r'τυπος παροχης', col, re.IGNORECASE):
                            type_value = normalize_text(row[columns[i]])
                            if pd.isna(row[columns[i]]) or (isinstance(row[columns[i]], str) and not type_value.startswith('γ')):
                                found = True
                                break
                    if not found:
                        match = False
                elif condition_type == 'None' or normalized_condition == '':
                    pass  # Fallback handled below
                else:
                    match = False

                conditions_met = match

            if conditions_met:
                logging.debug(f"Applied rule for {provider}, table {table_id}: {condition} -> {eidos_value}")
                return eidos_value

        errors.append(f"No matching rule for {provider}, table {table_id}, columns: {columns}, row: {row.to_dict()}")
        return ""
    except Exception as e:
        errors.append(f"Error setting Είδος for row {row.name} in provider {provider}, table {table_id}: {str(e)}")
        return ""

# Συνάρτηση για καθαρισμό αριθμητικής τιμής
def clean_numeric(value):
    if pd.isna(value):
        return pd.NA
    str_val = str(value).strip()
    str_val = re.sub(r'[^\d,.-]', '', str_val)
    # Try English format: , as thousands, . as decimal
    eng_val = str_val.replace(',', '')
    try:
        return float(eng_val)
    except ValueError:
        pass
    # Try Greek: . as thousands, , as decimal
    gr_val = str_val.replace('.', '').replace(',', '.')
    try:
        return float(gr_val)
    except ValueError:
        return pd.NA

# Συνάρτηση για καθαρισμό και mapping DataFrame
def clean_and_map_df(df, provider_key, provider_name, errors, sheet_name=None, table_title=None):
    try:
        df.columns = [str(col).strip() if pd.notna(col) else "" for col in df.columns]

        seen = {}
        new_cols = []
        for col in df.columns:
            if col == "Αρ. Παροχής":
                if col in seen:
                    new_cols.append("Αρ. Παροχής.1")
                else:
                    new_cols.append(col)
                seen[col] = 1
            else:
                new_cols.append(col)
        df.columns = new_cols

        # Καθαρισμός numeric στηλών πριν το set_eidos
        numeric_cols_patterns = [
            r'καθαρα', r'φπα', r'αμοιβη', r'commission', r'προμηθεια', r'kwh', r'καταναλωση', r'ενεργεια'
        ]
        for col in df.columns:
            if any(re.search(pattern, normalize_text(col), re.IGNORECASE) for pattern in numeric_cols_patterns):
                df[col] = df[col].apply(clean_numeric)

        table_id = get_table_id(df.columns, provider_key, sheet_name, valid_sheets, table_title)
        df["Είδος"] = df.apply(lambda row: set_eidos(row, provider_name, provider_key, table_id, errors, df.columns, eidos_rules, sheet_name, table_title), axis=1)

        def is_valid_row(row):
            if all(pd.isna(cell) or str(cell).strip() == "" for cell in row):
                logging.debug("Αποκλείστηκε κενή γραμμή")
                return False
            for i, cell in enumerate(row):
                cell_value = normalize_text(cell)
                if any(kw.lower() in cell_value for kw in ignore_keywords) and not re.search(r'αμοιβη|προμηθεια|commission', normalize_text(row.name), re.IGNORECASE):
                    logging.debug(f"Αποκλείστηκε γραμμή με τιμή: {cell_value}")
                    return False
            return True

        df = df[df.apply(is_valid_row, axis=1)]
        df = df.replace("01/01/1970", "")

        if provider_key not in mapping_dict:
            errors.append(f"No mapping found for provider_key {provider_key}")
            return pd.DataFrame()

        mapped_cols = mapping_dict[provider_key]
        df = df[[col for col in df.columns if col in mapped_cols or col == "Είδος"]]
        df = df.rename(columns=mapped_cols)

        for col in df.columns:
            if col not in new_column_names and col != "Είδος":
                errors.append(f"Unmapped column '{col}' in provider {provider_name}")

        allowed_cols = [col for col in new_column_names if col in df.columns] + (["Είδος"] if "Είδος" in df.columns else [])
        df = df[allowed_cols]
        
        # Αφού γίνει η χαρτογράφηση και προσθήκη "Πάροχος"
        df = df[df.apply(lambda row: row.notna().sum() >= 4, axis=1)]

        def parse_date(value):
            if pd.isna(value):
                return value
            if isinstance(value, datetime):
                return value.strftime('%d/%m/%Y')
            str_val = str(value).strip()
            for fmt in date_formats_df['Format']:
                try:
                    return datetime.strptime(str_val, fmt).strftime('%d/%m/%Y')
                except ValueError:
                    continue
            try:
                return pd.to_datetime(str_val, errors='coerce').strftime('%d/%m/%Y')
            except:
                return value

        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(parse_date)

        df.insert(0, "Πάροχος", provider_name)
        if normalize_text(table_title) == 'cl' or normalize_text(table_title) == 'claw back':
            if 'Καθαρά' in df.columns:
                df['Καθαρά'] *= -1
            if 'ΦΠΑ' in df.columns:
                df['ΦΠΑ'] *= -1
            if 'Αμοιβή' in df.columns:
                df['Αμοιβή'] *= -1
            else:
                # Αν ΦΠΑ λείπει, υπολόγισε αρνητικά μετά
                df['ΦΠΑ'] = df['Καθαρά'] * fpa_rate  # fpa_rate από config
                df['Αμοιβή'] = df['Καθαρά'] + df['ΦΠΑ']
        return df
    except Exception as e:
        errors.append(f"Error in clean_and_map_df for provider {provider_name}, sheet {sheet_name}: {str(e)}")
        return pd.DataFrame()

# Συνάρτηση για εμφάνιση παραθύρου εισαγωγής ημερομηνίας
def get_date_input():
    root = tk.Tk()
    root.title("Εισαγωγή Ημερομηνίας")
    root.geometry("300x150")
    tk.Label(root, text="Εισάγετε ημερομηνία (YYYY-MM):").pack(pady=10)
    date_entry = tk.Entry(root)
    date_entry.pack(pady=10)

    def submit():
        date_str = date_entry.get().strip()
        if validate_date(date_str):
            root.date = date_str
            root.destroy()
        else:
            messagebox.showerror("Σφάλμα", "Μη έγκυρη ημερομηνία. Χρησιμοποιήστε μορφή YYYY-MM (π.χ., 2025-05).")

    tk.Button(root, text="Υποβολή", command=submit).pack(pady=10)
    root.mainloop()
    return getattr(root, 'date', None)

# Ζήτα από τον χρήστη την ημερομηνία
date_input = get_date_input()
if not date_input:
    print("❌ Δεν εισήχθη ημερομηνία. Το πρόγραμμα τερματίστηκε.")
    exit()

# Επεξεργασία αρχείων
all_dfs = []
errors = []

for file in os.listdir():
    if file.endswith(".xlsx") and file != 'mapping_ενεργειας.xlsx' and file != output_file % date_input:
        file_clean = normalize_text(file)
        provider = next((row['Provider'] for _, row in file_mapping_df.iterrows() if re.search(row['FilePattern'].lower(), file_clean)), None)
        if not provider:
            errors.append(f"No provider matched for file {file}")
            continue
        try:
            xls = pd.ExcelFile(file, engine="openpyxl")
            for sheet in xls.sheet_names:
                if provider in valid_sheets and sheet not in valid_sheets[provider]:
                    errors.append(f"Skipping sheet {sheet} in file {file}: not in valid sheets for {provider}")
                    continue
                sheet_provider = "NLS_ΗΡΩΝ" if sheet == "ΗΡΩΝ" and provider == "NRG" else provider
                try:
                    df = xls.parse(sheet, header=None)
                    if df.empty or df.dropna(how='all').empty:
                        errors.append(f"Sheet {sheet} in file {file} is empty or contains no valid data")
                        continue
                    rows = df.values.tolist()
                    i = 0
                    table_title = None
                    while i < len(rows):
                        row = rows[i]
                        num_non_na = sum([1 for cell in row if pd.notna(cell)])
                        if num_non_na <= 1:
                            if any(isinstance(cell, str) and cell.strip() for cell in row):
                                table_title = next((cell for cell in row if isinstance(cell, str) and cell.strip()), None)
                            i += 1
                            continue
                        is_header, provider_key = is_header_row(row, sheet_provider, sheet)
                        if is_header:
                            header = [str(cell).strip() if pd.notna(cell) else "" for cell in row]
                            data = []
                            i += 1
                            while i < len(rows):
                                next_row = rows[i]
                                num_non_na_next = sum([1 for cell in next_row if pd.notna(cell)])
                                if num_non_na_next <= 1:
                                    if any(isinstance(cell, str) and cell.strip() for cell in next_row):
                                        break
                                    i += 1
                                    continue
                                next_is_header, _ = is_header_row(next_row, sheet_provider, sheet)
                                if next_is_header:
                                    break
                                data.append(next_row)
                                i += 1
                            temp_df = pd.DataFrame(data, columns=header)
                            cleaned_df = clean_and_map_df(temp_df, provider_key, sheet_provider, errors, sheet, table_title)
                            if not cleaned_df.empty:
                                errors.append(f"Processed table for provider {sheet_provider}, sheet: {sheet}, table_id: {get_table_id(header, provider_key, sheet, valid_sheets, table_title)}, columns: {list(temp_df.columns)}")
                                all_dfs.append(cleaned_df)
                        else:
                            i += 1
                except Exception as e:
                    errors.append(f"Error processing sheet {sheet} in file {file}: {str(e)}")
                    continue
        except Exception as e:
            errors.append(f"Error processing file {file}: {str(e)}")

# Συγχώνευση και εξαγωγή
valid_dfs = [df for df in all_dfs if not df.empty and not df.isna().all().all()]
if valid_dfs:
    final_df = pd.concat(valid_dfs, ignore_index=True)
    date_value = get_last_day_of_month(date_input)
    final_df.insert(0, "date", date_value)

    if 'Partner' in final_df.columns:
        final_df['Partner'] = final_df['Partner'].apply(lambda x: clean_partner(x, errors, partner_rules_df))

    if 'Είδος' in final_df.columns and 'Partner' in final_df.columns:
        cols = final_df.columns.tolist()
        partner_idx = cols.index('Partner')
        eidos = final_df.pop('Είδος')
        final_df.insert(partner_idx, 'Είδος', eidos)

    if 'Καθαρά' in final_df.columns and 'ΦΠΑ' in final_df.columns and 'Αμοιβή' in final_df.columns:
        final_df['Καθαρά'] = pd.to_numeric(final_df['Καθαρά'], errors='coerce')
        final_df['ΦΠΑ'] = pd.to_numeric(final_df['ΦΠΑ'], errors='coerce')
        final_df['ΦΠΑ'] = final_df.apply(
            lambda row: row['Καθαρά'] * fpa_rate if pd.isna(row['ΦΠΑ']) or row['ΦΠΑ'] == '' else row['ΦΠΑ'],
            axis=1
        )
        final_df['Αμοιβή'] = final_df['Καθαρά'] + final_df['ΦΠΑ']
        # Κάνε αρνητικά για Terminated
        if 'Κατάσταση Παροχής' in final_df.columns:
            mask = final_df['Κατάσταση Παροχής'] == 'Terminated'
            final_df.loc[mask, 'Καθαρά'] *= -1
            final_df.loc[mask, 'ΦΠΑ'] *= -1
            final_df.loc[mask, 'Αμοιβή'] *= -1

    writer = pd.ExcelWriter(output_file % date_input, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    final_df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    date_format = workbook.add_format({'num_format': 'yyyy-mm'})
    worksheet.set_column(0, 0, 12, date_format)
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    col_indices = {col: idx for idx, col in enumerate(final_df.columns)}
    for col_name in ['Καθαρά', 'ΦΠΑ', 'Αμοιβή']:
        if col_name in col_indices:
            worksheet.set_column(col_indices[col_name], col_indices[col_name], 10, number_format)
    writer.close()
    print(f"✅ Το αρχείο δημιουργήθηκε: {output_file % date_input}")
else:
    print("❌ Δεν βρέθηκαν έγκυρα δεδομένα για συγχώνευση.")

if errors:
    print("\n⚠️ Σφάλματα:")
    for err in errors:
        print("-", err)