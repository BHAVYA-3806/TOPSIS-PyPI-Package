# TOPSIS-BHAVYA-102203806

*A Python package for implementing the TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution) method.*

---

## 📖 Description

TOPSIS is a multi-criteria decision-making method that ranks alternatives based on their proximity to the ideal solution and distance from the worst solution. This package simplifies the computation, allowing users to rank options effectively by specifying criteria, weights, and impacts.

---

## 🛠 Installation

You can install the package via PyPI:  

[https://pypi.org/project/topsis-BHAVYA-102203806/](https://pypi.org/project/topsis-BHAVYA-102203806/)

```bash
pip install topsis-BHAVYA-102203806
```

Or, clone this repository and install the required dependencies:
bash
git clone https://github.com/bmahajan230/topsis-BHAVYA-102203806.git
cd topsis-BHAVYA-102203806
pip install -r requirements.txt


---

## 🚀 Usage

Run the package via the command line using the following syntax:

bash
python <program.py> <InputDataFile> "<Weights>" "<Impacts>" <ResultFileName>


### Example Command:
bash
python 102203806.py 102203806-data.csv "1,1,1,2" "+,+,-,+" 102203806-result.csv


### Input Parameters:
1. **<InputDataFile>**: Path to the input .csv file.
2. **<Weights>**: Comma-separated weights (e.g., "1,1,1,2").
3. **<Impacts>**: Comma-separated impacts (e.g., " +,+,-,+").
4. **<ResultFileName>**: Path to save the output .csv file.

---

## 📋 Input File Format

- *File Type*: .csv (Comma-Separated Values).
- *Columns*:
  - The *first column* should contain the names of the alternatives (e.g., M1, M2, M3).
  - Columns from the *2nd to last* must contain numeric values only.

### Example Input File (102203806-data.csv):

| Object | Criterion 1 | Criterion 2 | Criterion 3 | Criterion 4 |
|--------|-------------|-------------|-------------|-------------|
| M1     | 50          | 30          | 20          | 40          |
| M2     | 60          | 20          | 40          | 30          |

---

## 📤 Output File Format

The output .csv file will include the input data with two additional columns:
- **Topsis Score**: A numerical value indicating the relative closeness to the ideal solution.
- **Rank**: The rank of each alternative based on the score.

### Example Output File (102203806-result.csv):

| Object | Criterion 1 | Criterion 2 | Criterion 3 | Criterion 4 | Topsis Score | Rank |
|--------|-------------|-------------|-------------|-------------|--------------|------|
| M1     | 50          | 30          | 20          | 40          | 0.67         | 2    |
| M2     | 60          | 20          | 40          | 30          | 0.89         | 1    |

---

## 🧰 Features

1. *Error Handling*:
   - Ensures the correct number of parameters are provided.
   - Validates the input file format and contents.
   - Checks for non-numeric values in the criteria columns.
2. *Simple CLI Interface*:
   - Intuitive and easy to run from the command line.
3. *Customizable*:
   - Accepts user-defined weights and impacts for the criteria.

---

## ✅ Requirements

- Python 3.6 or above
- Required libraries:
  - pandas
  - numpy

Install dependencies using:
bash
pip install -r requirements.txt


---

## 📚 References

- [Learn the Mathematics of TOPSIS](https://en.wikipedia.org/wiki/TOPSIS)
- [How to Upload Python Package to PyPI](https://medium.com/@joel.barmettler/how-to-upload-your-python-package-to-pypi-65edc5fe9c56)

---

## 👩‍💻 Author

- *Name:* Bhavya  
- *Roll Number:* 102203806  
- *Email:* [bmahajan230@gmail.com](mailto:bmahajan230@gmail.com)  
- *GitHub:* [BHAVYA-3806](https://github.com/BHAVYA-3806)

---

## 📝 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

