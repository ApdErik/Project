"""
This script processes data from a CSV file, performs various statistical analyses,
and generates visualizations.

1. Imports Libraries: Imports necessary libraries for data manipulation, plotting,
    and database operations.
2. Read and Process Data:
    - `load_data(filepath)`: Loads data from a CSV file into a DataFrame
    - `process_data(report)`: Processes the DataFrame to compute statistics and summaries by region
3. Save Data to Excel:
    - `save_to_excel(regions, sorted_regions, regions_literacy)`: Saves processed data to an Excel
        file with multiple sheets.
4. Plotting Functions:
    - `plot_gdp_max_min(regions)`: Creates a bar chart of GDP max and min by
        region.
    - `plot_violin_gdp(sorted_regions, regions1)`: Creates violin plots of GDP distribution
        by region.
    - `plot_boxplots(sorted_regions, regions1)`: Creates boxplots of GDP distribution by region.
5. Load Data to SQL:
    - `load_to_sql(df, table_name)`: Loads a DataFrame into a SQL Server table.
6. Main Function:
    - Executes the data processing pipeline, including loading data, processing data,
        saving to Excel, generating plots, and loading data to SQL.

Usage:
    Ensure that the CSV file "nations.csv" is in the working directory.
    Run the script to process the data and generate the outputs.

Dependencies:
    - pandas
    - numpy
    - matplotlib
    - sqlalchemy
    - openpyxl
    - xlsxwriter
"""
# 1. Import Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sqlalchemy import create_engine

# 2. Read and Process Data
def load_data(filepath):
    """Load data from a CSV file

    Args:
        filepath (str): The path to the CSV file to be read

    Returns:
        pd.DataFrame: A Pandas DataFrame containing the data from the CSV file.
    """
    return pd.read_csv(filepath, encoding="ISO-8859-1", index_col=0)

def process_data(report):
    """Process the data to compute various statistics and summaries

    Args:
        report (pd.DataFrame): A DataFrame containing the data with columns 'region', 'gdp',
        'country', and 'literacy'.

    Returns:
        tuple: A tuple containing:
        - pd.DataFrame: A DataFrame with summary statistics for each region, including:
            - Number of countries
            - Maximum GDP
            - Minimum GDP
            - Country with the maximum GDP
            - Country with the minimum GDP
        - list: A list of unique regions.
        - pd.DataFrame: The original DataFrame(unchanged).
        - pd.DataFrame: A DataFrame with average literacy rates by region
    """
    regions = report["region"].value_counts().to_frame()
    regions.reset_index(drop=False, inplace=True)

    regions1 = report["region"].unique()
    gdp_max = {}
    gdp_min = {}
    country_max = {}
    country_min = {}

    for region in regions1:
        gdp_max_value = report[report["region"] == region]["gdp"].max()
        gdp_min_value = report[report["region"] == region]["gdp"].min()
        gdp_max[region] = gdp_max_value
        gdp_min[region] = gdp_min_value
        country_max[region] = report[report["gdp"] == gdp_max_value]["country"].values[0]
        country_min[region] = report[report["gdp"] == gdp_min_value]["country"].values[0]

    regions = pd.DataFrame({
        "No_Countries": report["region"].value_counts(),
        "gdp_max": pd.Series(gdp_max),
        "gdp_min": pd.Series(gdp_min),
        "country_max": pd.Series(country_max),
        "country_min": pd.Series(country_min)
    }).round({'gdp_max': 2, 'gdp_min': 2})

    regions.reset_index(drop=False, inplace=True)
    regions.rename(columns={"index": "Region"}, inplace=True)

    # Compute average literacy by region
    literacy_avg = {}
    for region in regions1:
        literacy_avg[region] = report[report["region"] == region]["literacy"].mean()

    regions_literacy = pd.DataFrame({
        "literacy_avg": pd.Series(literacy_avg)
    })
    regions_literacy.reset_index(drop=False, inplace=True)
    regions_literacy.rename(columns={"index": "Region"}, inplace=True)
    regions_literacy = regions_literacy.round({"literacy_avg": 2})

    return regions, regions1, report, regions_literacy

def save_to_excel(regions, sorted_regions, regions_literacy):
    """Save data to an Excel file with multiple sheets.

    Args:
        regions (dp.DataFrame): DataFrame containing summary statistics for each region,
        including numbers of countries, maximum and minimum GDP, and countries with
        maximum and minimum GDP.
        sorted_regions (dict): Dictionary where keys are region names and values are DataFrames
        containing data sorted by GDP.
        regions_literacy (pd.DataFrame): DataFrame containing average literacy rates for each region
    Returns:
        None
    """
    file_path = "sorted_regions_output.xlsx"

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        regions.to_excel(writer, sheet_name="Report", index=False)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for region, df in sorted_regions.items():
            df.iloc[:,2:13] = df.iloc[:,2:13].round(2)
            df.to_excel(writer, sheet_name=region, index=False)

        combined_df = pd.concat(sorted_regions.values(), ignore_index=True)
        combined_df.iloc[:, 2:13] = combined_df.iloc[:, 2:13].round(2)
        combined_df.to_excel(writer, sheet_name="continent_data_sorted", index=False)

        regions_literacy.to_excel(writer, sheet_name="Avg_Literacy_by_Region", index=False)

def plot_gdp_max_min(regions):
    """Create a bar chart showing the maximum and minimum GDP for each region.

    Args:
        regions (pd.DataFrame): DataFrame containing GDP maximum and minimum values by region.
        It should include the following columns: 'Region', 'gdp_max', and 'gdp_min'.
    Returns:
        None
    """
    x = np.arange(len(regions["Region"]))
    width = 0.35

    fig, ax = plt.subplots(figsize=(10, 6))
    bars1 = ax.bar(x - width / 2, regions["gdp_max"], width, label='GDP Max', color='skyblue')
    bars2 = ax.bar(x + width / 2, regions["gdp_min"], width, label='GDP Min', color='salmon')

    ax.set_xlabel('Region')
    ax.set_ylabel('GDP')
    ax.set_title('GDP Max and Min by Region')
    ax.set_xticks(x)
    ax.set_xticklabels(regions["Region"])
    ax.legend()

    # Add labels above the bars
    for bars in [bars1, bars2]:
        for ba in bars:
            height = ba.get_height()
            ax.text(
                ba.get_x() + ba.get_width() / 2,
                height,
                f'${height:.2f}',
                ha='center',
                va='bottom'
            )

    plt.tight_layout()
    plt.savefig('Bar_Charts_GDP_Region.png', dpi=300)
    plt.show()

def plot_violin_gdp(sorted_regions, regions1):
    """Create violin plots showing the distribution of GDP for each region.

    Args:
        sorted_regions (dict): Dictionary where keys are region names and values
        are DataFrames containing GDP data sorted by GDP.
        regions1 (list): List of region names.
    Returns:
        None
    """
    plt.style.use('ggplot')
    plt.figure(figsize=(10, 6))

    for i, region in enumerate(regions1):
        gdp_data = sorted_regions[region]["gdp"].dropna()
        plt.violinplot(gdp_data, positions=[i], showmeans=True, showmedians=False)
        median = np.median(gdp_data)
        plt.scatter(i, median, color='red', marker='o', label="Median" if i == 0 else None)

    plt.xticks(range(len(regions1)), regions1, rotation=45, ha="right")
    plt.ylabel("GDP")
    plt.title("Violin Plots of GDP by Region")
    plt.legend()
    plt.tight_layout()
    plt.savefig('ViolinPlot_GDP_Region.png', dpi=300)
    plt.show()

def plot_boxplots(sorted_regions, regions1):
    """Create boxplots showing the distribution of GDP for each region.

    Args:
        sorted_regions (dict): Dictionary where keys are region names and
        values are DataFrames containing GDP data sorted by GDP.
        regions1 (list): List of region names.
    Returns:
        None
    """
    plt.style.use('ggplot')
    plt.figure(figsize=(12, 8))

    data = [sorted_regions[region]["gdp"].dropna() for region in regions1]
    colormap = plt.get_cmap('Set2')
    colors = colormap(np.linspace(0, 1, len(regions1)))

    boxplots = plt.boxplot(data, positions=range(len(regions1)), patch_artist=True, showmeans=False)

    for i, patch in enumerate(boxplots['boxes']):
        patch.set_facecolor(colors[i])

    plt.xticks(range(len(regions1)), regions1, rotation=45, ha="right")
    plt.ylabel("GDP")
    plt.title("Boxplots of GDP by Region")
    plt.tight_layout()
    plt.savefig('Boxplots_GDP_Region.png', dpi=300)
    plt.show()

def load_to_sql(df, table_name):
    """Load a DataFrame into a SQL Server table.

    Args:
        df (pd.DataFrame): DataFrame to be loaded into the SQL table.
        table_name (str): Name of the SQL table where the DataFrame will be loaded.
    Returns:
        None
    """
    server = "DESKTOP-2P8QS6B"
    database = "Nations"
    driver = "ODBC Driver 17 for SQL Server"
    trusted_connection = "yes"
    connection_string = f'mssql+pyodbc://@{server}/{database}?driver={driver}&trusted_connection={trusted_connection}'
    engine = create_engine(connection_string)
    df.to_sql(table_name, con=engine, if_exists="replace", index=False)

def main():
    """Main function to execute the data processing pipeline.

    This function performs the following tasks:
    1. Loads data from a CSV file using 'load_data'.
    2. Processes the data to extract and calculate various metrics by region using 'process_data'.
    3. Sorts the data by GDP within each region.
    4. Saves the processed data to an Excel file using 'save_to_excel'.
    5. Generates and saves plots:
        - Bar chart of GDP max and min by region using 'plot_gdp_max_min'.
        - Violin plot of GDP distribution by regions using 'plot_violin_gdp'.
        - Boxplots of GDP distribution by region using 'plot_boxplots'.
    6. Combines the sorted regional data into a singel DataFrame and rounds the values.
    7. Loads the combined data and the average literacy data to SQL Server using 'load_to_sql'.
    """
    # Load and process data
    report = load_data("nations.csv")
    regions, regions1, report, regions_literacy = process_data(report)

    sorted_regions = {}
    for region in regions1:
        sorted_regions[region] = report[report["region"] == region].sort_values(by="gdp", ascending=False).reset_index(drop=True)

    # Save data to Excel
    save_to_excel(regions, sorted_regions, regions_literacy)

    # Generate plots
    plot_gdp_max_min(regions)
    plot_violin_gdp(sorted_regions, regions1)
    plot_boxplots(sorted_regions, regions1)

    # Load data to SQL
    combined_df = pd.concat(sorted_regions.values(), ignore_index=True)
    combined_df.iloc[:, 2:13] = combined_df.iloc[:, 2:13].round(2)
    load_to_sql(combined_df, "continent_data_sorted")
    load_to_sql(regions_literacy, "region_avg_literacy")

if __name__ == "__main__":
    main()
