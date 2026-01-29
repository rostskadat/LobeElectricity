# LobeElectricity

Generate an Excel report to be imported in Google Sheet for further analysis.

## How to use

1. Check the [default.yaml](default.yaml)
1. Load the environment and execute the script

    ```shell
    . .venv/bin/activate
    python extract_bill_information.py
    ```

1. Once the finished, import the resulting XSL file (`--output`) in Google Sheet: `File` >> `Import` >> `Replace spreadsheet`
1. Once imported you can apply the AppScript that will create all the required graphics `Extentions` >> `App Script` >> `Run`
