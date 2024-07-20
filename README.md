# Excel File Converter

This Python script converts Excel files from the `.xls` format to `.xlsx` format. It uses the `win32com.client` module to interact with Microsoft Excel.

## Table of Contents

- [Overview](#overview)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Overview

The script performs the following tasks:
1. Defines paths for the old version (`.xls`) and new version (`.xlsx`) directories.
2. Ensures that the new version directory exists.
3. Initializes an Excel application.
4. Iterates through all files in the old version directory.
5. Converts any `.xls` or `.xlsx` files to the `.xlsx` format and saves them in the new version directory.
6. Closes the Excel application after the process is complete.

## Requirements

- Python 3.x
- `pywin32` (Python for Windows extensions)

## Installation

1. **Clone the repository:**

    ```sh
    git clone https://github.com/yourusername/repository-name.git
    ```

2. **Navigate to the project directory:**

    ```sh
    cd repository-name
    ```

3. **Install the required Python packages:**

    Ensure you have `pywin32` installed. You can install it using `pip`:

    ```sh
    pip install pywin32
    ```

## Usage

1. **Prepare your directory structure:**

    Ensure you have two directories:
    - `old version` (containing `.xls` or `.xlsx` files)
    - `new version` (where converted `.xlsx` files will be saved)

2. **Run the script:**

    Execute the script using Python:

    ```sh
    python script_name.py
    ```

    Replace `script_name.py` with the name of your script file.


## Contributing

Contributions are welcome! If you have suggestions or improvements, please open an issue or submit a pull request.

1. **Fork the repository.**
2. **Create a new branch:**

    ```sh
    git checkout -b feature/your-feature-name
    ```

3. **Commit your changes:**

    ```sh
    git commit -am 'Add some feature'
    ```

4. **Push to the branch:**

    ```sh
    git push origin feature/your-feature-name
    ```

5. **Open a pull request.**

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
