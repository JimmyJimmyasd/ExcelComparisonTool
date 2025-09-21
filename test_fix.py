import sys
sys.path.append('.')

from analysis.statistical_analysis import StatisticalAnalyzer
from analysis.visualization import StatisticalVisualizer
import pandas as pd
import numpy as np

# Create simple test data
test_data = {
    'col1': [1, 2, 3, 4, 5],
    'col2': [2, 4, 6, 8, 10],
    'col3': [1, 3, 5, 7, 9]
}
df = pd.DataFrame(test_data)

# Test the visualization
try:
    analyzer = StatisticalAnalyzer()
    visualizer = StatisticalVisualizer()
    
    correlation_data = analyzer.calculate_correlation_matrix(df)
    print("Correlation data keys:", correlation_data.keys())
    
    if 'message' not in correlation_data:
        fig = visualizer.create_correlation_heatmap(correlation_data, 'pearson')
        print("✅ Correlation heatmap created successfully!")
    else:
        print("❌ No correlation data available:", correlation_data['message'])
        
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()