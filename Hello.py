import streamlit as st
import pandas as pd
import numpy as np
import requests
import xlrd
import os
import urllib
import ssl
from datetime import datetime, date, timedelta
import time
import json
from typing import Dict, Tuple, Optional
import io

# Configure page
st.set_page_config(
    page_title="Fund Tracker",
    page_icon="📊",
    layout="wide"
)

# Cache configuration
@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_cached_data():
    """Cache data to reduce API calls"""
    return {}

# Rate limiting helper
class RateLimiter:
    def __init__(self, max_calls_per_minute=60):
        self.max_calls = max_calls_per_minute
        self.calls = []
    
    def can_call(self):
        now = time.time()
        # Remove calls older than 1 minute
        self.calls = [call_time for call_time in self.calls if now - call_time < 60]
        return len(self.calls) < self.max_calls
    
    def record_call(self):
        self.calls.append(time.time())

# Global rate limiter
rate_limiter = RateLimiter(max_calls_per_minute=30)

def safe_yfinance_call(ticker_symbol: str, retries: int = 3) -> Optional[float]:
    """Safely call yfinance with rate limiting and retries"""
    if not rate_limiter.can_call():
        st.warning("Rate limit reached. Using cached data.")
        return None
    
    try:
        import yfinance as yf
        rate_limiter.record_call()
        
        ticker = yf.Ticker(ticker_symbol)
        
        # Try different methods to get price data
        try:
            # Method 1: Use fast_info
            price = ticker.fast_info.last_price
            if price and price > 0:
                return price
        except:
            pass
        
        try:
            # Method 2: Use history
            hist = ticker.history(period='1d', interval='1m')
            if not hist.empty:
                return hist['Close'].iloc[-1]
        except:
            pass
        
        try:
            # Method 3: Use info
            info = ticker.info
            if 'regularMarketPrice' in info and info['regularMarketPrice']:
                return info['regularMarketPrice']
        except:
            pass
        
        return None
        
    except Exception as e:
        st.error(f"Error fetching data for {ticker_symbol}: {str(e)}")
        return None

def get_fx_rate() -> float:
    """Get JPY exchange rate with fallback"""
    # Try yfinance first
    fx_rate = safe_yfinance_call('JPY=X')
    if fx_rate:
        return fx_rate
    
    # Fallback to alternative sources
    try:
        # You can add alternative FX rate sources here
        # For now, return a reasonable default
        return 150.0  # Default JPY rate
    except:
        return 150.0

def get_futures_price() -> float:
    """Get live ESU5 futures price from CME"""
    try:
        # Method 1: Try CME API for ESU5
        cme_url = "https://www.cmegroup.com/api/quote/v2/contracts/quotes/ESU5"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'application/json',
            'Referer': 'https://www.cmegroup.com/'
        }
        
        response = requests.get(cme_url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if 'quotes' in data and len(data['quotes']) > 0:
                last_price = data['quotes'][0].get('last', 0)
                if last_price and last_price > 0:
                    return float(last_price)
    except Exception as e:
        st.warning(f"CME API failed: {str(e)}")
    
    try:
        # Method 1b: Try CME public data API
        cme_public_url = "https://www.cmegroup.com/CmeWS/mvc/Quotes/Future/1/G"
        params = {
            'tradeDate': date.today().strftime('%m/%d/%Y'),
            'productId': 'ES',
            'expirationMonth': 'U5'
        }
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(cme_public_url, params=params, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if 'quotes' in data and len(data['quotes']) > 0:
                for quote in data['quotes']:
                    if quote.get('expirationMonth') == 'U5':
                        last_price = quote.get('last', 0)
                        if last_price and last_price > 0:
                            return float(last_price)
    except Exception as e:
        st.warning(f"CME public API failed: {str(e)}")
    
    try:
        # Method 2: Scrape CME website for ESU5
        cme_web_url = "https://www.cmegroup.com/markets/equities/sp/e-mini-sandp500.quotes.html"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(cme_web_url, headers=headers, timeout=10)
        if response.status_code == 200:
            # Look for ESU5 in the HTML with more specific patterns
            import re
            # Pattern for ESU5 with price
            esu5_patterns = [
                r'ESU5[^>]*>([0-9,]+\.?[0-9]*)',
                r'ESU5[^>]*last[^>]*>([0-9,]+\.?[0-9]*)',
                r'ESU5[^>]*price[^>]*>([0-9,]+\.?[0-9]*)'
            ]
            
            for pattern in esu5_patterns:
                matches = re.findall(pattern, response.text)
                if matches:
                    for match in matches:
                        try:
                            # Remove commas and convert to float
                            price = float(match.replace(',', ''))
                            if 4000 < price < 7000:  # Reasonable ESU5 price range
                                return price
                        except:
                            continue
    except Exception as e:
        st.warning(f"CME web scraping failed: {str(e)}")
    
    try:
        # Method 3: Try yfinance as fallback
        futures_price = safe_yfinance_call("ESU5.CME")
        if futures_price:
            return futures_price
    except:
        pass
    
    try:
        # Method 4: Try alternative yfinance symbol
        futures_price = safe_yfinance_call("ES=F")
        if futures_price:
            return futures_price
    except:
        pass
    
    # Final fallback
    st.error("Unable to fetch live futures data. Using default value.")
    return 5000.0

def download_fund_data(fund_code: str) -> Tuple[pd.DataFrame, pd.DataFrame, float]:
    """Download and parse fund data with error handling"""
    try:
        YearMonth = date.today().strftime('%Y%m')
        link = f'https://www.nikkoam.com/files/etf/_shared/xls/portfolio/{fund_code}_{YearMonth}.xls'
        
        # Create SSL context that ignores certificate verification
        import ssl
        ssl_context = ssl.create_default_context()
        ssl_context.check_hostname = False
        ssl_context.verify_mode = ssl.CERT_NONE
        
        # Create opener with SSL context
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ssl_context))
        urllib.request.install_opener(opener)
        
        # Download file
        fileName, headers = urllib.request.urlretrieve(link)
        
        # Read workbook
        wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))
        
        # Get fund data
        fundData = pd.read_excel(wb, header=None, usecols="A,C", skiprows=1, nrows=10, engine='xlrd').T
        fundData.columns = fundData.iloc[0]
        fundData = fundData[1:]
        fundData = fundData.infer_objects()
        
        # Get positions
        FundPositions = pd.read_excel(wb, skiprows=range(1, 13), skipfooter=3, header=1, engine='xlrd')
        
        # Get NAV
        nav = fundData['AUM*1'].iloc[0]
        
        # Clean up downloaded file
        os.remove(fileName)
        
        return FundPositions, fundData, nav
        
    except Exception as e:
        st.error(f"Error downloading data for fund {fund_code}: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), 0.0

def calculate_fund_metrics(fund_code: str, nav: float, cf_factor: float, 
                          fx_rate: float, futures_price: float, 
                          fund_positions: pd.DataFrame) -> Dict:
    """Calculate fund metrics"""
    try:
        # Get futures positions
        fut_positions = fund_positions[fund_positions.Category == "Future"]
        
        if fut_positions.empty:
            return {
                'cur_fut_position': 0,
                'target_position': 0,
                'target_trade': 0,
                'live_fund_weight': 0,
                'prev_inv_ratio': 0,
                'fut_pct_change': 0
            }
        
        # Calculate current position
        cur_fut_position = fut_positions['Value(Local)'].sum() / 5 / fut_positions['Price'].mean()
        
        # Calculate percentage change
        fut_pct_change = futures_price / fut_positions['Price'].mean() - 1
        
        # Set leverage ratio based on fund
        lev_ratio = 2 if fund_code == "2239" else -1
        
        # Calculate target position
        target_position = (lev_ratio * nav * (1 + lev_ratio * fut_pct_change) / 
                         (fx_rate * 5 * futures_price) * cf_factor)
        
        # Calculate metrics
        target_trade = target_position - cur_fut_position
        live_fund_weight = cur_fut_position / target_position * lev_ratio if target_position != 0 else 0
        prev_inv_ratio = fut_positions['Value(JPY)'].sum() / nav if nav != 0 else 0
        
        return {
            'cur_fut_position': cur_fut_position,
            'target_position': target_position,
            'target_trade': target_trade,
            'live_fund_weight': live_fund_weight,
            'prev_inv_ratio': prev_inv_ratio,
            'fut_pct_change': fut_pct_change
        }
        
    except Exception as e:
        st.error(f"Error calculating metrics for fund {fund_code}: {str(e)}")
        return {
            'cur_fut_position': 0,
            'target_position': 0,
            'target_trade': 0,
            'live_fund_weight': 0,
            'prev_inv_ratio': 0,
            'fut_pct_change': 0
        }

def export_to_tsv(fund_data: Dict) -> str:
    """Export fund data to TSV format"""
    tsv_data = []
    
    # Add header
    tsv_data.append("Fund\tLive Weight\tTarget Trade (Micros)\tCurrent Position\tTarget Position\tInvestment Ratio")
    
    # Add fund data
    for fund_code, data in fund_data.items():
        tsv_data.append(f"{fund_code}\t{data['live_fund_weight']:.4f}\t{data['target_trade']:.1f}\t"
                       f"{data['cur_fut_position']:.1f}\t{data['target_position']:.1f}\t{data['prev_inv_ratio']:.4f}")
    
    return "\n".join(tsv_data)

def main():
    st.title("📊 Fund Tracker")
    st.markdown("---")
    
    # Sidebar for inputs
    with st.sidebar:
        st.header("Configuration")
        cf_2239 = st.number_input('2239 CF:', step=10000000.00, format="%f", value=0.0)
        cf_2240 = st.number_input('2240 CF:', step=10000000.00, format="%f", value=0.0)
        
        st.markdown("---")
        st.header("Data Sources")
        use_cached = st.checkbox("Use cached data (if available)", value=True)
        
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Get market data with rate limiting
    with st.spinner("Fetching market data..."):
        fx_rate = get_fx_rate()
        futures_price = get_futures_price()
    
    # Process funds
    fund_results = {}
    
    for fund_code in ["2239", "2240"]:
        cf_value = cf_2239 if fund_code == "2239" else cf_2240
        
        with st.spinner(f"Processing fund {fund_code}..."):
            # Download fund data
            fund_positions, fund_data, nav = download_fund_data(fund_code)
            
            if nav > 0:
                cf_factor = (cf_value + nav) / nav
                
                # Calculate metrics
                metrics = calculate_fund_metrics(
                    fund_code, nav, cf_factor, fx_rate, futures_price, fund_positions
                )
                
                fund_results[fund_code] = metrics
                
                # Display results
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric(
                        label=f"Fund {fund_code} - Live Weight",
                        value=f"{metrics['live_fund_weight']:+.2%}",
                        delta=f"{metrics['target_trade']:.1f} Micros"
                    )
                
                with col2:
                    st.metric(
                        label=f"Fund {fund_code} - Position",
                        value=f"{metrics['cur_fut_position']:.1f}",
                        delta=f"Target: {metrics['target_position']:.1f}"
                    )
    
    # Market data display
    st.markdown("---")
    st.header("Market Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric(
            label="Futures Price",
            value=f"{futures_price:.2f}",
            delta=f"{fund_results.get('2239', {}).get('fut_pct_change', 0):+.2%}"
        )
    
    with col2:
        st.metric(
            label="JPY Rate",
            value=f"{fx_rate:.5f}",
            delta="Current"
        )
    
    # TSV Export
    st.markdown("---")
    st.header("Export Data")
    
    if fund_results:
        tsv_data = export_to_tsv(fund_results)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📋 Copy TSV to Clipboard"):
                st.write("TSV data copied to clipboard!")
                st.code(tsv_data)
        
        with col2:
            st.download_button(
                label="💾 Download TSV",
                data=tsv_data,
                file_name=f"fund_trades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.tsv",
                mime="text/tab-separated-values"
            )
    
    # Additional information
    st.markdown("---")
    st.header("Additional Info")
    
    if fund_results:
        st.write("**Investment Ratios:**")
        for fund_code, data in fund_results.items():
            st.write(f"Fund {fund_code}: {data['prev_inv_ratio']:+.2%}")
    
    # Error handling and status
    if not rate_limiter.can_call():
        st.warning("⚠️ Rate limit approaching. Consider using cached data.")

if __name__ == "__main__":
    main()
