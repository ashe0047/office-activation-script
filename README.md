# Office Activation Python Script

![Python](https://img.shields.io/badge/Python-3.6%20%7C%203.7%20%7C%203.8%20%7C%203.9-blue.svg)
![License](https://img.shields.io/github/license/yourusername/office-activation-script)

## Description

This is a Python script designed to automate the activation process for Microsoft Office products. It utilizes the `ospp.vbs` script provided by Microsoft to activate Office using a valid product key or perform other activation-related tasks.

The script can be used with different Office versions and licensing types, including retail and volume licenses (KMS, MAK). It allows you to activate Office installations offline using confirmation IDs or online using Microsoft's activation servers.

## Features

- Activate Microsoft Office using product keys.
- Uninstall an existing product key from Office installations.
- Perform offline activation using confirmation IDs.
- Check license information for installed Office products.
- Reset licensing status for Office installations.
- And more!

## Requirements

- Python 3.6 or above.
- Windows operating system.
- A valid Microsoft Office product key (if using activation features).

## Installation

1. Clone this repository to your local machine:

  ```bash
   git clone https://github.com/yourusername/office-activation-script.git
  ```
2. Navigate to the project directory:
   
  ```bash
  cd office-activation-script
  ```

3. Install the required Python packages:

  ```bash
  pip install -r requirements.txt
  ```
## Usage

To activate Microsoft Office using the script, run the following command:
  ```bash
  python activate_office.py --action activate --product-key YOUR_PRODUCT_KEY
  ```
Replace `YOUR_PRODUCT_KEY` with your valid Office product key.

For more options and features, refer to the documentation in the `docs` directory.

## Documentation

For detailed instructions, options, and usage examples, please refer to the documentation.

## License

This project is licensed under the MIT License.
