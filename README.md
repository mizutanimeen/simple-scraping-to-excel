# Simple Scraping To Excel

## Outline

Given the URL and tags, run the scraping and save all of its text in cell An (n=1,2,3,,) of the `output.xlsx` file.

After saving, the `output.xlsx` can be automatically opened in Excel.

## Set Up

### chrome driver
1. check the chrome version from [here](chrome://settings/help)
2. download the chrome driver from [here](https://chromedriver.chromium.org/downloads) and put it in your project.
### .env
Create `.env` with reference to `example.env`

* URL
    * Enter the URL you wish to access
* DRIVER_PATH
    * The path where you put the chrome driver should be relative to main.py.
* TAG_NAME
    * Enter the tag you wish to retrieve
    * e.g. h3, div, etc...
* IS_OPEN
    * After saving the xlsx file, enter `True` if you want to open it.

## Execute.
```
python main.py
```


