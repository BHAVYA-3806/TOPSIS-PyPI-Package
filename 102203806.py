import sys
import os
import pandas as pd
import numpy as np

def conversion_to_csv(inp_file, roll_no):
    try:
        # Check if the input Excel file exists
        if not os.path.exists(inp_file):
            raise FileNotFoundError(f"File '{inp_file}' not found.")

        # Read the Excel file
        data = pd.read_excel(inp_file)

        # Generate the new CSV file name
        new_file_csv = f"{roll_no}-data.csv"

        # Save the data as a CSV file
        data.to_csv(new_file_csv, index=False)
        print(f"The input file convertion is successful and is saved as '{new_file_csv}'.")
        return new_file_csv

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


def topsis(inp_file_csv, weights, impacts, out_file_csv):
    try:
        data = pd.read_csv(inp_file_csv)

        if len(data.columns) < 3:
            raise ValueError("Input file must have three or more columns.")

        mat = data.iloc[:, 1:].values
        if not np.issubdtype(mat.dtype, np.number):
            raise ValueError("From 2nd to last columns must contain numeric values only.")

        weights = [float(w) for w in weights.split(',')]
        impacts = impacts.split(',')

        if len(weights) != mat.shape[1] or len(impacts) != mat.shape[1]:
            raise ValueError("No. of weights, impacts, and columns (from 2nd to last) should be same.")

        if not all(i in ['+', '-'] for i in impacts):
            raise ValueError("Impacts must be either '+' or '-'.")

        norm_mat = mat / np.sqrt((mat ** 2).sum(axis=0))

        weighted_mat = norm_mat * weights

        ideal_best = []
        ideal_worst = []
        for i in range(len(impacts)):
            if impacts[i] == '+':
                ideal_best.append(np.max(weighted_mat[:, i]))
                ideal_worst.append(np.min(weighted_mat[:, i]))
            else:
                ideal_best.append(np.min(weighted_mat[:, i]))
                ideal_worst.append(np.max(weighted_mat[:, i]))

        ideal_best = np.array(ideal_best)
        ideal_worst = np.array(ideal_worst)

        best_dist = np.sqrt(((weighted_mat - ideal_best) ** 2).sum(axis=1))
        worst_dist = np.sqrt(((weighted_mat - ideal_worst) ** 2).sum(axis=1))

        topsis_score = worst_dist / (best_dist + worst_dist)

        data['Topsis Score'] = topsis_score
        data['Rank'] = data['Topsis Score'].rank(ascending=False).astype(int)

        out_file = f"{102203806}-result.csv"
        data.to_csv(out_file, index=False)
        print(f"TOPSIS Analysis completed successfully and the results are saved to {out_file}.")

    except FileNotFoundError:
        print(f"Error: File '{inp_csv}' not found.")
    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    inp_file = "data.xlsx"
    roll_no = "102203806"
    weights = '0.2,0.3,0.3,0.1,0.1'
    impacts = '+,+,-,+,-'

     # Convert and rename Excel to CSV
    inp_csv = conversion_to_csv(inp_file, roll_no)

    topsis(inp_csv, weights, impacts, roll_no)