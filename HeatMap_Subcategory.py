import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap

# Load the spreadsheet data
file_path = '/Users/nicolashock/Documents/GitHub/CME-Volume-Insights/CME_FIA_ETD_Tracker.xlsx'
data = pd.read_excel(file_path)

# Create the custom colormap
start_color = '#081d37'  # Represents the lowest volume
end_color = '#04e273'    # Represents the highest volume
custom_cmap = LinearSegmentedColormap.from_list("custom_cmap", [start_color, end_color])

# Define the function to plot and save heat maps for Subcategory
def plot_heatmap_for_venue_subcategory(data, venue=None):
    if venue:
        # Filter data for the specified venue
        filtered_data = data[data['Venue'] == venue]
        title = f'Volume Composition by Subcategory for {venue} (2022 vs 2023) in Percentage'
        filename = f'heatmap_subcategory_{venue}.png'
    else:
        # Use all data for the total CME Group Exchanges heat map
        filtered_data = data
        title = 'Total Volume Composition by Subcategory for CME Group Exchanges (2022 vs 2023) in Percentage'
        filename = 'heatmap_subcategory_total.png'

    # Pivot the data for Subcategory
    pivot_data = filtered_data.pivot_table(values='Volume', index='Subcategory', columns='Year', aggfunc='sum')

    # Calculate percentages for each Subcategory
    pivot_data_percentage = pivot_data.div(pivot_data.sum(axis=0), axis=1) * 100

    # Plotting the heat map using the custom colormap for Subcategory
    plt.figure(figsize=(10, 6))
    sns.heatmap(pivot_data_percentage, cmap=custom_cmap, annot=True, fmt=".1f", cbar_kws={'label': 'Percentage of Total Volume'})
    plt.title(title)
    plt.ylabel('Subcategory')
    plt.xlabel('Year')
    plt.tight_layout()
    
    # Save the figure with the Subcategory filename
    full_path = f'/Users/nicolashock/Documents/GitHub/CME-Volume-Insights/{filename}'
    plt.savefig(full_path, format='png', dpi=300)
    
    # Close the figure to free up memory
    plt.close()

# Plot heat maps for each individual venue and the total for Subcategory
venues = ['CME', 'CBOT', 'NYMEX', 'COMEX', None]  # Include None for the total CME Group Exchanges
for venue in venues:
    plot_heatmap_for_venue_subcategory(data, venue=venue)
