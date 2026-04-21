import re

def convert_eur_to_standard_format(value_str: str, prefer_decimal: bool = False) -> str:
    if not value_str or not isinstance(value_str, str):
        return value_str
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        dot_pos = value_str.rfind('.')
        comma_pos = value_str.rfind(',')
        if comma_pos > dot_pos:
            converted = value_str.replace('.', '').replace(',', '.')
        else:
            return value_str
        try:
            num = float(converted)
            if '.' in converted:
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            return f"{num:,.0f}"
        except ValueError:
            return value_str
    elif ',' in value_str and '.' not in value_str:
        converted = value_str.replace(',', '.')
        try:
            num = float(converted)
            decimal_places = len(converted.split('.')[-1])
            return f"{num:,.{decimal_places}f}"
        except ValueError:
            return value_str
    elif '.' in value_str:
        parts_list = value_str.split('.')
        if not prefer_decimal and len(parts_list) == 2 and len(parts_list[1]) == 3:
            converted = value_str.replace('.', '')
            try:
                num = float(converted)
                return f"{num:,.0f}"
            except ValueError:
                return value_str
        else:
            try:
                num = float(value_str)
                decimal_places = len(value_str.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
    else:
        try:
            num = float(value_str)
            return f"{num:,.0f}"
        except ValueError:
            return value_str

# Test cases
test_weight = "1.155"
print(f"Original: {test_weight}")
print(f"Default (VW): {convert_eur_to_standard_format(test_weight, prefer_decimal=False)}")
print(f"Skoda: {convert_eur_to_standard_format(test_weight, prefer_decimal=True)}")

test_price = "89.03"
print(f"Original: {test_price}")
print(f"Skoda Price: {convert_eur_to_standard_format(test_price, prefer_decimal=True)}")

test_thousands = "1.000"
print(f"Original: {test_thousands}")
print(f"Skoda Thousand/Decimal: {convert_eur_to_standard_format(test_thousands, prefer_decimal=True)}")
