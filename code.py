import kagglehub
import os
import pandas as pd
from tkinter import *
from tkinter import ttk
from openpyxl import Workbook
import tkinter.font as tk_font


def hide_buttons_in_region(x_min, x_max, y_min, y_max):
    """Hides all buttons within a specific screen region"""
    for widget in window.winfo_children():
        if widget.grid_info():
            x, y = widget.grid_info()["column"], widget.grid_info()["row"]
            if x_min <= x <= x_max and y_min <= y <= y_max:
                widget.grid_forget()


def on_selection_year(event):
    """Display buttons based on year selection"""
    hide_buttons_in_region(0, 2, 2, 3)  # hide previous selection
    if choice_1.get() == '1 Year':
        label_4.grid(row=2, column=0)
        choice_3.grid(row=3, column=0)

    elif choice_1.get() == 'Multiple Years':
        choice_4.grid(row=3, column=0)
        choice_5.grid(row=3, column=2)
        label_5.grid(row=2, column=0)
        label_6.grid(row=2, column=2)
        label_7.grid(row=3, column=1)


def on_selection_regions(event):
    """Display buttons based on region selection"""
    hide_buttons_in_region(3, 3, 2, 3)  # hide previous selection
    if choice_2.get() == 'Countries':
        label_8.grid(row=2, column=3)
        choice_6.grid(row=3, column=3)

    elif choice_2.get() == 'Regions':
        label_9.grid(row=2, column=3)
        choice_7.grid(row=3, column=3)


def year_check():
    """Check if valid year or year range used"""
    years_or_year = 'none'
    year_set = []
    if choice_1.get() == 'Multiple Years':  # check which choice made
        years_or_year = 'Year Range'
        if start_year.get() == end_year.get():  # same year for year range used not valid
            years_or_year = 'none'
        else:
            try:
                if int(start_year.get()) not in years:  # checks if inside year set in database
                    years_or_year = 'none'
            except ValueError:
                years_or_year = 'none'
            try:
                if int(end_year.get()) not in years:
                    years_or_year = 'none'
            except ValueError:
                years_or_year = 'none'
        if years_or_year == 'Year Range':  # gets year range if passes all tests
            first_year = min(int(start_year.get()), int(end_year.get()))
            last_year = max(int(start_year.get()), int(end_year.get()))
            for i in range(first_year, last_year + 1):
                year_set.append(i)
    elif choice_1.get() == '1 Year':  # 1 year selected
        years_or_year = 'Year'
        try:
            if int(one_year.get()) not in years:
                years_or_year = 'none'
            else:
                year_set.append(int(one_year.get()))  #add year to year set
        except ValueError:
            years_or_year = 'none'

    return years_or_year, year_set


def area_check():
    """Check if valid based on region/country selected"""
    region_or_country = 'none'
    area_set = ''
    if choice_2.get() == 'Regions':  # check based on choice
        region_or_country = 'Region'
        try:
            if interest_region.get() not in regions:
                region_or_country = 'none'
            else:
                area_set = interest_region.get()  # get region
        except ValueError:
            region_or_country = 'none'
    elif choice_2.get() == 'Countries':
        region_or_country = 'Country'
        try:
            if interest_country.get() not in countries:
                region_or_country = 'none'
            else:
                area_set = interest_country.get()  # get country
        except ValueError:
            region_or_country = 'none'
    return region_or_country, area_set


def region_filter(year_set, area_set):
    """Get set of countries in particular regions"""
    if area_set == 'Africa':
        country_set = ['Algeria', 'Angola', 'Benin', 'Botswana', 'Burkina Faso', 'Burundi', 'Cabo Verde', 'Cameroon',
                       'Central African Republic', 'Chad', 'Comoros', 'Democratic Republic of Congo', 'Djibouti',
                       'Egypt',
                       'Equatorial Guinea', 'Eritrea', 'Eswatini', 'Ethiopia', 'Gabon', 'Gambia', 'Ghana', 'Guinea',
                       'Guinea-Bissau',
                       'Ivory Coast', 'Kenya', 'Lesotho', 'Liberia', 'Libya', 'Madagascar', 'Malawi', 'Mali',
                       'Mauritania',
                       'Mauritius', 'Morocco', 'Mozambique', 'Namibia', 'Niger', 'Nigeria', 'Rwanda',
                       'Sao Tome and Principe',
                       'Senegal', 'Seychelles', 'Sierra Leone', 'Somalia', 'South Africa', 'South Sudan', 'Sudan',
                       'Tanzania', 'Togo',
                       'Tunisia', 'Uganda', 'Zambia', 'Zimbabwe']
    elif area_set == 'Asia':
        country_set = ['Afghanistan', 'Armenia', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Bhutan', 'Brunei', 'Cambodia',
                       'China',
                       'Christmas Island', 'Cyprus', 'Georgia', 'Hong Kong', 'India', 'Indonesia', 'Iran', 'Iraq',
                       'Israel', 'Japan',
                       'Jordan', 'Kazakhstan', 'Kuwait', 'Kyrgyzstan', 'Laos', 'Lebanon', 'Macao', 'Malaysia',
                       'Maldives', 'Mongolia',
                       'Myanmar', 'Nepal', 'North Korea', 'Oman', 'Pakistan', 'Palestine', 'Philippines', 'Qatar',
                       'Saudi Arabia', 'Singapore',
                       'South Korea', 'Sri Lanka', 'Syria', 'Taiwan', 'Tajikistan', 'Thailand', 'Timor-Leste', 'Turkey',
                       'Turkmenistan',
                       'United Arab Emirates', 'Uzbekistan', 'Vietnam', 'Yemen']
    elif area_set == 'Europe':
        country_set = ['Albania', 'Andorra', 'Austria', 'Belarus', 'Belgium', 'Bosnia and Herzegovina', 'Bulgaria',
                       'Croatia', 'Czechia',
                       'Denmark', 'Estonia', 'Faroe Islands', 'Finland', 'France', 'Germany', 'Greece', 'Greenland',
                       'Hungary', 'Iceland',
                       'Ireland', 'Italy', 'Kosovo', 'Latvia', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Malta',
                       'Moldova', 'Monaco', 'Montenegro',
                       'Netherlands', 'North Macedonia', 'Norway', 'Poland', 'Portugal', 'Romania', 'Russia',
                       'San Marino', 'Serbia', 'Slovakia',
                       'Slovenia', 'Spain', 'Sweden', 'Switzerland', 'Ukraine', 'United Kingdom']
    elif area_set == 'North America':
        country_set = ['Antigua and Barbuda', 'Bahamas', 'Barbados', 'Belize', 'Bermuda', 'Canada', 'Costa Rica',
                       'Cuba', 'Dominica',
                       'Dominican Republic', 'El Salvador', 'Greenland', 'Grenada', 'Guatemala', 'Haiti', 'Honduras',
                       'Jamaica', 'Mexico',
                       'Nicaragua', 'Panama', 'Saint Kitts and Nevis', 'Saint Lucia',
                       'Saint Vincent and the Grenadines',
                       'Trinidad and Tobago', 'United States']
    elif area_set == 'Oceania':
        country_set = ['Australia', 'Fiji', 'Kiribati', 'Marshall Islands', 'Micronesia (country)', 'Nauru',
                       'New Caledonia',
                       'New Zealand', 'Palau', 'Papua New Guinea', 'Samoa', 'Solomon Islands', 'Tonga', 'Tuvalu',
                       'Vanuatu',
                       'Wallis and Futuna']
    elif area_set == 'South America':
        country_set = ['Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Ecuador', 'Guyana', 'Paraguay', 'Peru',
                       'Suriname', 'Uruguay', 'Venezuela']
    else:
        country_set = []
    region_filtered_df = df[(df['Year'].isin(year_set)) & (df['Entity'].isin(country_set))]
    # filter data to country in region and in year range
    return region_filtered_df


def co2_unit_write(number):
    """Change value to suitable carbon emission unit"""
    if abs(number) >= 1e9:  # Giga
        number = number / 1e9
        unit = " Gt CO\u2082"
    elif abs(number) >= 1e6:  # Mega
        number = number / 1e6
        unit = " Mt CO\u2082"
    elif abs(number) >= 1e3:  # Kilo
        number = number / 1e3
        unit = " kt CO\u2082"
    else:
        unit = " t CO\u2082"
    if number <= 0:
        output_string = str(round(number, 2)) + unit
    else:
        output_string = '+' + str(round(number, 2)) + unit  #indicates positive growth
    return output_string


def get_max_width(data, col_index, font):
    """Returns the maximum pixel width of a column based on the data."""
    return max(font.measure(str(row[col_index])) for row in data) + 10  # add padding


def analysis_calculate(years_or_year, year_set, region_or_country, area_set):
    """Calculate necessary parameters and store them"""
    filtered_df = df[(df['Year'].isin(year_set)) & (df['Entity'] == area_set)]  # filter of year and region
    region_filtered_df = region_filter(year_set, area_set)  # country filter region
    year_group_filtered_df = filtered_df.groupby('Year')['growth_emissions_total'].sum()  # group filtered data by year
    # group filtered data by region
    country_group_region_filtered_df = region_filtered_df.groupby('Entity')['growth_emissions_total'].sum()
    if len(year_set) > 1:  # convert year range to strings for display
        year_display = str(min(year_set)) + ' - ' + str(max(year_set))
    else:
        year_display = str(year_set[0])
    total_value = float(filtered_df['growth_emissions_total'].sum())
    total_value_string = co2_unit_write(total_value)
    # common parameters and details
    parameters = [{"Parameter": region_or_country, "Value": area_set},
                  {"Parameter": years_or_year, "Value": year_display},
                  {"Parameter": 'Total Growth Emissions', "Value": total_value_string}]
    if years_or_year == 'Year Range' and region_or_country == 'Region':  # parameters based on category selected
        average_annual = co2_unit_write(float(year_group_filtered_df.mean()))
        parameters.append({"Parameter": 'Average Annual Carbon Growth Emissions:',
                           "Value": average_annual})
        average_national = co2_unit_write(float(country_group_region_filtered_df.mean()))
        parameters.append({"Parameter": 'Average National Carbon Growth Emissions:',
                           "Value": average_national})
        answer = co2_unit_write(float(year_group_filtered_df.max()))
        max_year = year_group_filtered_df.idxmax()
        max_annual = str(max_year) + ', ' + answer
        parameters.append({"Parameter": 'Max Annual Carbon Growth Emissions:',
                           "Value": max_annual})
        answer = co2_unit_write(float(year_group_filtered_df.min()))
        min_year = year_group_filtered_df.idxmin()
        min_annual = str(min_year) + ', ' + answer
        parameters.append({"Parameter": 'Min Annual Carbon Growth Emissions',
                           "Value": min_annual})
        answer = co2_unit_write(float(country_group_region_filtered_df.max()))
        max_country = country_group_region_filtered_df.idxmax()
        max_national = str(max_country) + ', ' + answer
        parameters.append({"Parameter": 'Max National Carbon Growth Emissions',
                           "Value": max_national})
        answer = co2_unit_write(float(country_group_region_filtered_df.min()))
        min_country = country_group_region_filtered_df.idxmin()
        min_national = str(min_country) + ', ' + answer
        parameters.append({"Parameter": 'Min National Carbon Growth Emissions',
                           "Value": min_national})
    elif years_or_year == 'Year Range' and region_or_country == 'Country':
        average_annual = co2_unit_write(float(year_group_filtered_df.mean()))
        parameters.append({"Parameter": 'Average Annual Carbon Growth Emissions:',
                           "Value": average_annual})
        answer = co2_unit_write(float(year_group_filtered_df.max()))
        max_year = year_group_filtered_df.idxmax()
        max_annual = str(max_year) + ', ' + answer
        parameters.append({"Parameter": 'Max Annual Carbon Growth Emissions:',
                           "Value": max_annual})
        answer = co2_unit_write(float(year_group_filtered_df.min()))
        min_year = year_group_filtered_df.idxmin()
        min_annual = str(min_year) + ', ' + answer
        parameters.append({"Parameter": 'Min Annual Carbon Growth Emissions',
                           "Value": min_annual})
    elif years_or_year == 'Year' and region_or_country == 'Region':
        average_national = co2_unit_write(float(country_group_region_filtered_df.mean()))
        parameters.append({"Parameter": 'Average National Carbon Growth Emissions:',
                           "Value": average_national})
        answer = co2_unit_write(float(country_group_region_filtered_df.max()))
        max_country = country_group_region_filtered_df.idxmax()
        max_national = str(max_country) + ', ' + answer
        parameters.append({"Parameter": 'Max National Carbon Growth Emissions',
                           "Value": max_national})
        answer = co2_unit_write(float(country_group_region_filtered_df.min()))
        min_country = country_group_region_filtered_df.idxmin()
        min_national = str(min_country) + ', ' + answer
        parameters.append({"Parameter": 'Min National Carbon Growth Emissions',
                           "Value": min_national})
    width_para = get_max_width(parameters, 'Parameter', tk_font.Font())
    width_val = get_max_width(parameters, 'Value', tk_font.Font())
    width = [width_para, width_val]  # get width for report table
    return parameters, width


def save_report(tree):
    """Save generated report"""
    data = [['Parameters', 'Values']]
    for item in tree.get_children(): # get values in table
        data.append(tree.item(item)["values"])
    wb = Workbook()
    ws = wb.active
    for row in data:
        ws.append(row) # write xlsx file
    # Save as Excel file
    output_string = 'Carbon_Growth_Emissions_'+ data[1][1] + '_' + data[2][0] + str(data[2][1]) + '.xlsx'
    wb.save(output_string)


def write_report():
    """Generate and write report on display window"""
    years_or_year, year_set = year_check()
    region_or_country, area_set = area_check()#check viable inputs used
    if years_or_year in ['Year', 'Year Range'] and region_or_country in ['Region', 'Country']:
        # calculate values for report
        parameters, width = analysis_calculate(years_or_year, year_set, region_or_country, area_set)
        hide_buttons_in_region(0, 4, 4, 4) # clear previous table
        #set table
        columns = ['Parameter', 'Value']
        tree = ttk.Treeview(window, columns=columns, show="headings")
        for col in range(len(columns)):
            tree.heading(columns[col], text=columns[col])  # Set column headings
            tree.column(columns[col], width=width[col])
        # write table
        for i in range(len(parameters)):
            table_row = [parameters[i]["Parameter"], parameters[i]["Value"]]
            tree.insert("", END, values=table_row)
        tree.grid(row=4, column=0)
        button_2 = Button(window, text='Save')
        button_2.config(command=lambda: save_report(tree))
        button_2.grid(row=5, column=4)


# Code
# Download dataset
path = kagglehub.dataset_download("samithsachidanandan/year-on-year-change-in-co-emissions")
csv_files = [f for f in os.listdir(path) if f.endswith(".csv")]
df = pd.read_csv(os.path.join(path, csv_files[0]))
# list of countries and years
countries = sorted(df['Entity'].unique())
regions = ['Africa', 'Asia', 'Europe', 'North America', 'Oceania', 'South America']
regions_excluded = ['Asia (excl. China and India)', 'Europe (excl. EU-27)', 'Europe (excl. EU-28)',
                    'European Union (27)', 'European Union (28)', 'High-income countries',
                    'International aviation', 'International shipping', 'Low-income countries',
                    'Lower-middle-income countries', 'North America (excl. USA)',
                    'Upper-middle-income countries', 'World']
# exclude regions and unnecessary regions from country list
for reg in regions:
    countries.remove(reg)
for reg_excl in regions_excluded:
    countries.remove(reg_excl)
# year list
years = sorted(df['Year'].unique(), reverse=True)
# make display board
window = Tk()
window.title('Carbon Growth Emissions')
# variables
year_range_sel = StringVar()
start_year = StringVar()
end_year = StringVar()
one_year = StringVar()
country_region = StringVar()
interest_region = StringVar()
interest_country = StringVar()
# labels
label_1 = Label(window, text='Select Year/Years:')
label_1.grid(row=0, column=0)
label_2 = Label(window, text='Region:')
label_2.grid(row=0, column=3)
label_4 = Label(window, text='Select Year:')
label_5 = Label(window, text='From Year:')
label_6 = Label(window, text='To Year:')
label_7 = Label(window, text='-')
label_8 = Label(window, text='Select Country:')
label_9 = Label(window, text='Select Region:')
# buttons
button_1 = Button(window, text='Enter')
button_1.config(command=write_report)
button_1.grid(row=3, column=4)
# choice buttons
choice_1 = ttk.Combobox(window, values=['1 Year', 'Multiple Years'], textvariable=year_range_sel)
choice_1.grid(row=1, column=0)
choice_1.bind("<<ComboboxSelected>>", on_selection_year)
choice_2 = ttk.Combobox(window, values=['Countries', 'Regions'], textvariable=country_region)
choice_2.grid(row=1, column=3)
choice_2.bind("<<ComboboxSelected>>", on_selection_regions)
choice_3 = ttk.Combobox(window, values=years, textvariable=one_year)
choice_4 = ttk.Combobox(window, values=years, textvariable=start_year)
choice_5 = ttk.Combobox(window, values=years, textvariable=end_year)
choice_6 = ttk.Combobox(window, values=countries, textvariable=interest_country)
choice_7 = ttk.Combobox(window, values=regions, textvariable=interest_region)
window.mainloop()
