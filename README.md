Stock Market Analysis & Financial Dashboard

📌 Project Overview

This project is a Stock Market Analysis & Financial Dashboard that fetches real-time stock data, performs financial calculations, and generates visualizations. The data is stored in a database and exported to Excel with an embedded stock price chart.

🔹 Features

Real-Time Stock Data: Fetch stock prices using the Yahoo Finance API.

Data Storage: Store stock data in an SQLite database using SQLAlchemy.

Financial Metrics: Calculate key indicators like:

Daily Returns

50-day & 200-day Moving Averages

Sharpe Ratio (risk-adjusted return metric)

News Integration: Fetch recent stock-related news from a News API.

Data Visualization: Plot stock price trends with moving averages using Matplotlib & Seaborn.

Automated Excel Reports: Save stock data and insert stock price graph into Excel.

🛠 Technologies Used

Python: Core programming language.

Pandas & NumPy: Data analysis & calculations.

Matplotlib & Seaborn: Data visualization.

Yahoo Finance API: Fetch stock prices.

Requests & BeautifulSoup: API calls & web scraping.

SQLAlchemy (SQLite): Database management.

OpenPyXL: Automating Excel reports with graphs.

🚀 Installation & Setup

1️⃣ Clone the Repository

git clone https://github.com/PranayPrasanth/stock-analysis-dashboard.git
cd stock-analysis-dashboard

2️⃣ Install Required Libraries

pip install -r requirements.txt

3️⃣ Set Up Environment Variables

Create a .env file and add your API keys:

NEWS_API=your_news_api_key
NEWS_ENDPOINT=https://newsapi.org/v2/everything

4️⃣ Run the Project

python main.py



📈 Future Enhancements

Add more financial indicators like RSI & Bollinger Bands.

Develop a Flask/Streamlit web dashboard.

Implement automatic alerts for stock price changes.

Integrate Machine Learning models for stock trend prediction.

🔗 Contributing

Feel free to fork this repo and contribute! 🚀

📩 Contact

For questions, reach out via:
📧 Email: pranayprasanth200@gmail.com🔗 LinkedIn: https://www.linkedin.com/in/pranay-prasanth-97664a276/
