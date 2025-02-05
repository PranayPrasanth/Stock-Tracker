import pandas as pd
import numpy as np
import yfinance as yf
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine, MetaData
from datetime import datetime
from yahoo_fin import stock_info
import requests
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv
from yahoo_fin import stock_info
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl import Workbook


# Load API keys and endpoint from environment variables
load_dotenv()

NEWS_API = os.environ['NEWS_API']
NEWS_ENDPOINT = os.environ['NEWS_ENDPOINT']

TICKER = "AAPL"

# Fetch company name from Yahoo Finance, fallback to user input if not found
stock = yf.Ticker(TICKER)
COMPANY_NAME = stock.info.get('longName', "Company name not found")
if COMPANY_NAME == "Company name not found":
    COMPANY_NAME = input("Enter the company name: ")


def fetch_stock_data(ticker):
    """
    Fetch historical stock data for the given ticker using Yahoo Finance.
    
    Parameters:
        ticker (str): Stock ticker symbol.
    
    Returns:
        pd.DataFrame: Stock data including Open, High, Low, Close, Adj Close, and Volume.
    """
    today = datetime.today().date()
    base_year = f"{today.year - 5}-{today.month}-{today.day}"  # Fetch last 5 years of data
    stock = yf.download(ticker, start=base_year, end=today, auto_adjust=False)
    stock.reset_index(inplace=True)
    return stock


# Create SQLite database connection
engine = create_engine("sqlite:///finance.db")


def save_to_sql(dataframe, table_name):
    """
    Save a Pandas DataFrame to an SQLite database.

    Parameters:
        dataframe (pd.DataFrame): Data to save.
        table_name (str): Name of the table in the database.
    """
    dataframe.to_sql(table_name, engine, if_exists="replace", index=False)


def calculate_metrics(stock_data):
    """
    Calculate financial metrics such as daily returns and moving averages.

    Parameters:
        stock_data (pd.DataFrame): Stock data containing adjusted close prices.

    Returns:
        pd.DataFrame: Updated stock data with calculated metrics.
    """
    stock_data["Daily Return"] = stock_data["Adj Close"].pct_change()
    stock_data["50-Day MA"] = stock_data["Adj Close"].rolling(window=50).mean()
    stock_data["200-Day MA"] = stock_data["Adj Close"].rolling(window=200).mean()
    return stock_data


def plot_stock_data(stock_data, ticker, graph_filename="stock_chart.png"):
    """
    Plot stock price with 50-day and 200-day moving averages.

    Parameters:
        stock_data (pd.DataFrame): Stock data containing adjusted close prices.
        ticker (str): Stock ticker symbol.
    """
    plt.figure(figsize=(12, 6))
    plt.plot(stock_data["Date"], stock_data["Adj Close"], label="Adj Close", color="blue")
    plt.plot(stock_data["Date"], stock_data["50-Day MA"], label="50-Day MA", linestyle="dashed", color="red")
    plt.plot(stock_data["Date"], stock_data["200-Day MA"], label="200-Day MA", linestyle="dashed", color="green")
    plt.xlabel("Date")
    plt.ylabel("Stock Price")
    plt.title(f"{ticker} Stock Price & Moving Averages")
    plt.legend()
    plt.savefig(graph_filename)  # Save the plot as an image file
    plt.show()
    


def calculate_sharpe_ratio(stock_data, risk_free_rate=0.02):
    """
    Calculate the Sharpe Ratio to measure risk-adjusted returns.

    Parameters:
        stock_data (pd.DataFrame): Stock data with daily returns.
        risk_free_rate (float): Risk-free rate (default: 2%).

    Returns:
        float: Sharpe Ratio rounded to two decimal places.
    """
    avg_return = stock_data["Daily Return"].mean() * 252  # Annualized Return
    std_dev = stock_data["Daily Return"].std() * np.sqrt(252)  # Annualized Volatility
    sharpe_ratio = (avg_return - risk_free_rate) / std_dev
    return round(sharpe_ratio, 2)

def export_to_excel(stock_data, filename="financial_analysis.xlsx", graph_filename="stock_chart.png"):
    """
    Export stock data to an Excel file and add the stock price graph.
    
    Parameters:
        stock_data (pd.DataFrame): Data to export.
        filename (str): Output Excel file name (default: "financial_analysis.xlsx").
        graph_filename (str): Name of the graph image file to insert (default: "stock_chart.png").
    """
    # Flatten MultiIndex columns if they exist
    if isinstance(stock_data.columns, pd.MultiIndex):
        stock_data = stock_data.copy()
        stock_data.columns = ['_'.join(col).strip() for col in stock_data.columns.values]

    # Save DataFrame to Excel
    stock_data.to_excel(filename, index=False, engine="openpyxl")
    print(f"Data exported to {filename}")

    # Load the Excel file and get the active sheet
    wb = load_workbook(filename)
    ws = wb.active

    # Check if the graph file exists before inserting
    if os.path.exists(graph_filename):
        img = Image(graph_filename)
        ws.add_image(img, "H2")  # Insert the graph at cell H2
        wb.save(filename)
        print(f"Graph successfully added to {filename}")
    else:
        print(f"Graph file {graph_filename} not found. Make sure to save it before exporting.")


def get_news(ticker, api, endpoint):
    """
    Fetch the latest news articles related to a given stock ticker.

    Parameters:
        ticker (str): Stock ticker symbol.
        api (str): API key for the news service.
        endpoint (str): API endpoint for fetching news.

    Returns:
        list: Top 5 news articles as tuples (title, URL).
    """
    try:
        stock = yf.Ticker(ticker)
        company_name = stock.info.get('longName', "Company name not found")
    except Exception as e:
        return f"Error: {e}"

    news_parameters = {
        'apiKey': api,
        'q': company_name,
        'sortBy': 'relevancy',  # Sorting by relevancy, can be changed to 'popularity' or 'publishedAt'
        'language': 'en'  # English articles only
    }

    news_response = requests.get(url=endpoint, params=news_parameters)
    news_response.raise_for_status()  # Raise an error if the request fails
    news_data = news_response.json()

    articles = news_data.get('articles', [])  # Fetch articles safely

    top_5 = [(article['title'], article['url']) for article in articles[:5]]
    for index,article in enumerate(top_5):
        with open(f"{COMPANY_NAME} Top Articles.txt", 'a') as file:
            file.write(f"{index+ 1}. Title: {article[0]}\nLink: {article[1]}\n") # Display top 5 news articles

# Example Usages

stock_data = fetch_stock_data(TICKER) # Fetch Stock Data

save_to_sql(stock_data, f"{COMPANY_NAME}") # Save raw data to SQL
 
stock_data = calculate_metrics(stock_data) # Perform calculations

plot_stock_data(stock_data, TICKER) # Plot and save stock graph

calculate_sharpe_ratio(stock_data) # Calculate sharpe ratio

export_to_excel(stock_data) # Export the data to excel

get_news(TICKER, NEWS_API,NEWS_ENDPOINT) # Get Latest News about the Company



