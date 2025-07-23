from datetime import datetime, timedelta
import yfinance as yf

tickers = []
quantidade = int(input("Digite quantas ações serão adicionadas: "))

for i in range(quantidade):
    ticker = input(f"Digite a sigla da ação {i+1} que você quer analisar: ").upper()
    tickers.append(ticker)

rf = float(input("Valor do risk free (ex: 0.2 para 20%): "))

duracao = float(input("Digite o período de tempo que deseja analisar em anos (ex: 0.5 para 6 meses): ")) 

dias = int(duracao * 365)

end_date = datetime.today()
start_date = end_date - timedelta(days=dias)

start_date_str = start_date.strftime("%Y-%m-%d")
end_date_str = end_date.strftime("%Y-%m-%d")

df = yf.download(tickers, start_date, end_date, auto_adjust=True)

returns_dict = {}

for ticker in tickers:
    close_prices = df['Close'][ticker]
    returns = close_prices.pct_change().dropna()
    returns_dict[ticker] = returns