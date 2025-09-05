# -*- coding: utf-8 -*-
"""
Created on Fri Sep  5 19:48:16 2025

@author: saish
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ===============================
# Load and clean data
# ===============================
def load_data(file_path, drop_first=True, drop_last=True):
    df = pd.read_excel(file_path)
    if drop_first and drop_last:
        df = df.iloc[:, 1:-1]
    elif drop_first:
        df = df.iloc[:, 1:]
    elif drop_last:
        df = df.iloc[:, :-1]
    return df

# ===============================
# Helper: create output folders
# ===============================
def create_output_dirs(base_path):
    corr_dir = os.path.join(base_path, "correlation_matrices")
    freq_dir = os.path.join(base_path, "frequency_distributions")
    stats_dir = os.path.join(base_path, "statistical_summary")
    
    os.makedirs(corr_dir, exist_ok=True)
    os.makedirs(freq_dir, exist_ok=True)
    os.makedirs(stats_dir, exist_ok=True)
    
    return corr_dir, freq_dir, stats_dir

# ===============================
# Correlation Analysis
# ===============================
def correlation_analysis(df, output_dir, methods=["pearson", "kendall", "spearman"], n_cols=None):
    if n_cols:
        df = df.iloc[:, :n_cols]

    results = {}
    for method in methods:
        corr_matrix = df.corr(method=method)
        results[method] = corr_matrix

        # Save correlation matrix to Excel
        corr_file = os.path.join(output_dir, f"correlation_{method}.xlsx")
        corr_matrix.to_excel(corr_file)
        
        # Plot heatmap
        plt.figure(figsize=(10, 8))
        sns.heatmap(corr_matrix, annot=True, fmt=".2f", cmap="coolwarm", square=True)
        plt.title(f"Correlation Matrix ({method.capitalize()})")
        plt.savefig(os.path.join(output_dir, f"correlation_{method}.png"))
        plt.close()

    return results

# ===============================
# Frequency Distributions
# ===============================
def frequency_distributions(df, output_dir):
    numeric_cols = df.select_dtypes(include=np.number).columns
    for col in numeric_cols:
        plt.figure(figsize=(8, 5))
        sns.histplot(df[col], bins=20, kde=True, color="skyblue")
        plt.title(f"Frequency Distribution of {col}")
        plt.xlabel(col)
        plt.ylabel("Frequency")
        plt.savefig(os.path.join(output_dir, f"{col}_frequency.png"))
        plt.close()

# ===============================
# Statistical Summary
# ===============================
def statistical_summary(df, output_dir):
    summary_stats = pd.DataFrame()

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            col_data = df[col].dropna()
            mode_val = col_data.mode().values
            mode_val = mode_val[0] if len(mode_val) > 0 else np.nan

            summary_stats[col] = pd.Series({
                "Count": col_data.count(),
                "Mean": col_data.mean(),
                "Median": col_data.median(),
                "Mode": mode_val,
                "Std": col_data.std(),
                "Variance": col_data.var(),
                "Min": col_data.min(),
                "Max": col_data.max(),
                "25%": col_data.quantile(0.25),
                "50%": col_data.quantile(0.50),
                "75%": col_data.quantile(0.75),
                "Skewness": col_data.skew(),
                "Kurtosis": col_data.kurt()
            })

    summary_stats = summary_stats.T
    
    # Save to Excel
    stats_file = os.path.join(output_dir, "statistical_summary.xlsx")
    summary_stats.to_excel(stats_file)
    
    return summary_stats

# ===============================
# Main utility function
# ===============================
def run_analysis(file_path, n_cols=None, corr_methods=["pearson", "kendall", "spearman"],
                 do_correlation=True, do_frequency=True, do_stats=True):
    
    base_dir = os.path.dirname(os.path.abspath(file_path))
    corr_dir, freq_dir, stats_dir = create_output_dirs(base_dir)
    
    df = load_data(file_path)
    
    if do_correlation:
        correlation_analysis(df, corr_dir, methods=corr_methods, n_cols=n_cols)
    
    if do_frequency:
        frequency_distributions(df, freq_dir)
    
    if do_stats:
        statistical_summary(df, stats_dir)
    
    print("Analysis complete. Results saved in subfolders of:", base_dir)