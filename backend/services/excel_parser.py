import pandas as pd
import logging

logger = logging.getLogger(__name__)

def parse_excel_data(path):
    try:
        logger.debug(f"Reading Excel file: {path}")
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(path)
        data = {}
        
        for sheet_name in excel_file.sheet_names:
            logger.debug(f"Reading sheet: {sheet_name}")
            df = pd.read_excel(path, sheet_name=sheet_name)
            data[sheet_name] = df
            logger.debug(f"Successfully read sheet: {sheet_name}")
        
        return data
    except Exception as e:
        logger.error(f"Error parsing Excel file {path}: {str(e)}")
        raise

def get_sentiment_data(df, company_name):
    try:
        logger.debug(f"Getting sentiment data for company: {company_name}")
        # Filter data for specific company
        company_data = df[df['Company'] == company_name]
        
        if company_data.empty:
            logger.warning(f"No data found for company: {company_name}")
            return pd.DataFrame()
        
        # Group by date and get sentiment counts
        sentiment_data = company_data.groupby('Date')['Sentiment'].value_counts().unstack(fill_value=0)
        logger.debug("Successfully processed sentiment data")
        return sentiment_data
    except Exception as e:
        logger.error(f"Error getting sentiment data: {str(e)}")
        raise

def get_sentiment_counts(df):
    try:
        logger.debug("Getting sentiment counts")
        # Get total sentiment counts
        sentiment_counts = df['Sentiment'].value_counts()
        logger.debug("Successfully got sentiment counts")
        return sentiment_counts
    except Exception as e:
        logger.error(f"Error getting sentiment counts: {str(e)}")
        raise

def get_company_sentiment_counts(df):
    try:
        logger.debug("Getting company sentiment counts")
        # Get sentiment counts by company
        company_sentiments = df.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)
        logger.debug("Successfully got company sentiment counts")
        return company_sentiments
    except Exception as e:
        logger.error(f"Error getting company sentiment counts: {str(e)}")
        raise
