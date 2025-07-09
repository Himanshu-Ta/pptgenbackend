# This file contains the data for generating all native, editable charts.

# --- Data for Manufacturer Profile Chart ---
manufacturer_profile_data = [
    {'name': 'MAT-3', 'PriSeg1': 20, 'PriSeg2': 20, 'PriSeg3': 20, 'PriSeg4': 20, 'PriSeg5': 20},
    {'name': 'MAT-2', 'PriSeg1': 20, 'PriSeg2': 20, 'PriSeg3': 20, 'PriSeg4': 20, 'PriSeg5': 20},
    {'name': 'MAT-1', 'PriSeg1': 20, 'PriSeg2': 20, 'PriSeg3': 20, 'PriSeg4': 20, 'PriSeg5': 20},
    {'name': 'MAT',   'PriSeg1': 20, 'PriSeg2': 20, 'PriSeg3': 20, 'PriSeg4': 20, 'PriSeg5': 20},
    {'name': 'SPOT',  'PriSeg1': 20, 'PriSeg2': 20, 'PriSeg3': 20, 'PriSeg4': 20, 'PriSeg5': 20},
]

# --- Data for IMB Volume & Share Chart ---
imb_volume_data = [
    {'name': 'FY24', 'category1': 8.70, 'category2': 4.40},
    {'name': 'FY25', 'category1': 7.40, 'category2': 3.50},
    {'name': 'FY26', 'category1': 6.40, 'category2': 3.10},
    {'name': 'FY27', 'category1': 5.30, 'category2': 2.60},
    {'name': 'FY28', 'category1': 4.60, 'category2': 2.30},
    {'name': 'FY29', 'category1': 4.00, 'category2': 2.00},
    {'name': 'FY30', 'category1': 3.50, 'category2': 1.80},
]

imb_volume_share_data = [
    {'name': 'FY24', 'value': 40.0},
    {'name': 'FY25', 'value': 40.0},
    {'name': 'FY26', 'value': 40.0},
    {'name': 'FY27', 'value': 40.0},
    {'name': 'FY28', 'value': 40.0},
    {'name': 'FY29', 'value': 40.0},
    {'name': 'FY30', 'value': 42.0},
]

# --- Data for Consumer Typology (as a Table) ---
consumer_typology_data = [
    # Headers
    ["", "Total", "Typology 1", "Typology 2", "Typology 3", "Typology 4", "Typology 5", "Typology 6", "Typology 7"],
    # Data Rows
    ['Global', 'x', 10, 20, 10, 30, 10, 10, 10],
    ['Market', 'x', 10, 20, 10, 30, 10, 10, 10],
    ['Category 1', 'x', 10, 20, 20, 30, 10, 10, 10],
    ['Category 2', 'x', 10, 20, 20, 30, 10, 10, 10],
    ['Category 3', 'x', 10, 20, 10, 30, 20, 10, 10],
    ['Category 4', 'x', 10, 20, 10, 30, 10, 10, 20],
]

# --- Data for Price Ladder Chart ---
price_ladder_data = [
    {"series_name": "IMB Brand 1", "points": [(1, 4.0, 1.5), (1, 3.5, 2.5), (1, 3.8, 1.0)]},
    {"series_name": "IMB Brand 2", "points": [(2, 6.2, 0.8), (2, 5.2, 1.2), (2, 2.8, 2.0)]},
    {"series_name": "IMB Brand 3", "points": [(3, 4.5, 3.0), (3, 2.8, 1.8), (3, 3.2, 2.2)]},
    {"series_name": "IMB Brand 4", "points": [(4, 2.5, 2.5), (4, 2.2, 1.5), (4, 1.8, 3.5)]},
    {"series_name": "IMB Brand 5", "points": [(5, 6.4, 0.9), (5, 2.2, 1.1), (5, 1.9, 0.7)]},
    {"series_name": "IMB Brand 6", "points": [(6, 6.0, 1.0), (6, 2.0, 1.4), (6, 3.8, 0.6)]},
    {"series_name": "IMB Brand 7", "points": [(7, 6.2, 0.5), (7, 3.6, 1.3), (7, 2.5, 2.2)]}
]

price_segments = [
    {'name': 'PriSeg1 20.0%\nLC10.00 or more', 'y': 6.00},
    {'name': 'PriSeg2 20.0%\nLC8.00-9.99', 'y': 4.50},
    {'name': 'PriSeg3 20.0%\nLC6.00-7.99', 'y': 3.25},
    {'name': 'PriSeg4 20.0%\nLC3.00-5.99', 'y': 2.25},
    {'name': 'PriSeg5 20.0%\nLC2.99 or less', 'y': 1.00},
]

brand_summary_data = [
    {'brand': 'IMB Brand 1', 'ms': '5.00%', 'skus': 5, 'avgMs': '1.22%'},
    {'brand': 'IMB Brand 2', 'ms': '5.00%', 'skus': 8, 'avgMs': '1.38%'},
    {'brand': 'IMB Brand 3', 'ms': '5.00%', 'skus': 18, 'avgMs': '1.12%'},
    {'brand': 'IMB Brand 4', 'ms': '5.00%', 'skus': 7, 'avgMs': '0.48%'},
    {'brand': 'IMB Brand 5', 'ms': '5.00%', 'skus': 4, 'avgMs': '0.26%'},
    {'brand': 'IMB Brand 6', 'ms': '5.00%', 'skus': 11, 'avgMs': '0.40%'},
    {'brand': 'IMB Brand 7', 'ms': '5.00%', 'skus': 18, 'avgMs': '0.40%'},
]

brand_colors = ['d62728', 'ff7f0e', '1f77b4', '2ca02c', '9467bd', '8c564b', 'e377c2']
