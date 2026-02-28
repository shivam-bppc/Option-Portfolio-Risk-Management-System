import yfinance as yf
import pandas as pd
import numpy as np
from scipy.stats import norm
from scipy.optimize import fsolve
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
from nsepython import option_chain
from datetime import datetime
from scipy.interpolate import griddata

excel_filename = 'FRAM_Assignment_Results.xlsx'
writer = pd.ExcelWriter(excel_filename, engine='openpyxl')


print("\n" + "="*50)
print("PART A: DATA COLLECTION & STATISTICS")
print("="*50)

ticker_symbol = "ICICIBANK.NS"
stock = yf.Ticker(ticker_symbol)

print(f"Fetching data for {ticker_symbol}...")
hist_data = stock.history(period="3mo")

hist_data['Log_Returns'] = np.log(hist_data['Close'] / hist_data['Close'].shift(1))
log_returns = hist_data['Log_Returns'].dropna()

annualized_volatility = log_returns.std() * np.sqrt(252)
skewness = log_returns.skew()
kurtosis = log_returns.kurt()

S = hist_data['Close'].iloc[-1]

print(f"\nCurrent Stock Price (S): {S:.2f}")
print(f"Annualized Volatility (Sigma): {annualized_volatility:.4f}")
print(f"Skewness: {skewness:.4f}")
print(f"Kurtosis: {kurtosis:.4f}")

summary_stats = pd.DataFrame({
    'Metric': ['Current Stock Price', 'Annualized Volatility', 'Skewness', 'Kurtosis'],
    'Value': [S, annualized_volatility, skewness, kurtosis]
})
summary_stats.to_excel(writer, sheet_name='Summary Statistics', index=False)

hist_data_export = hist_data[['Close', 'Log_Returns']].copy()
hist_data_export.index = hist_data_export.index.tz_localize(None)
hist_data_export.to_excel(writer, sheet_name='Historical Data')

plt.figure(figsize=(10, 5))
plt.plot(log_returns)
plt.title(f'Daily Log Returns - {ticker_symbol}')
plt.xlabel('Date')
plt.ylabel('Log Return')
plt.grid(True)
plt.savefig('log_returns_plot.png', dpi=300, bbox_inches='tight')
plt.show()


print("\n" + "="*50)
print("PART B: OPTION PRICING (BSM MODEL)")
print("="*50)

def d1(S, K, T, r, sigma):
    return (np.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))

def d2(S, K, T, r, sigma):
    return d1(S, K, T, r, sigma) - sigma * np.sqrt(T)

def bsm_pricer(S, K, T, r, sigma, option_type):
    if option_type == 'call':
        price = S * norm.cdf(d1(S, K, T, r, sigma)) - K * np.exp(-r * T) * norm.cdf(d2(S, K, T, r, sigma))
    elif option_type == 'put':
        price = K * np.exp(-r * T) * norm.cdf(-d2(S, K, T, r, sigma)) - S * norm.cdf(-d1(S, K, T, r, sigma))
    return price

def iv_solver(sigma, S, K, T, r, market_price, option_type):
    bsm_price = bsm_pricer(S, K, T, r, sigma, option_type)
    return market_price - bsm_price

r = 0.07
K_atm = S
strikes = [K_atm * 0.95, K_atm * 0.98, K_atm, K_atm * 1.02, K_atm * 1.05]
maturities_days = [30, 60, 90]

pricing_results = []
for T_days in maturities_days:
    T_years = T_days / 365.0
    for K in strikes:
        call_price = bsm_pricer(S, K, T_years, r, annualized_volatility, 'call')
        put_price = bsm_pricer(S, K, T_years, r, annualized_volatility, 'put')
        pricing_results.append({
            'Maturity (Days)': T_days,
            'Strike': round(K, 2),
            'Call Price': round(call_price, 2),
            'Put Price': round(put_price, 2)
        })

pricing_table = pd.DataFrame(pricing_results)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print("\n--- BSM Option Pricing Table ---")
print(pricing_table.to_string(index=False))

pricing_table.to_excel(writer, sheet_name='Option Pricing', index=False)

print("\n" + "="*50)
print("PART C: GREEKS & VOLATILITY")
print("="*50)

def delta(S, K, T, r, sigma, option_type):
    d_1 = d1(S, K, T, r, sigma)
    return norm.cdf(d_1) if option_type == 'call' else norm.cdf(d_1) - 1

def gamma(S, K, T, r, sigma):
    return norm.pdf(d1(S, K, T, r, sigma)) / (S * sigma * np.sqrt(T))

def vega(S, K, T, r, sigma):
    return S * norm.pdf(d1(S, K, T, r, sigma)) * np.sqrt(T) * 0.01

def theta(S, K, T, r, sigma, option_type):
    d_1 = d1(S, K, T, r, sigma)
    d_2 = d2(S, K, T, r, sigma)
    term1 = - (S * norm.pdf(d_1) * sigma) / (2 * np.sqrt(T))
    if option_type == 'call':
        term2 = -r * K * np.exp(-r * T) * norm.cdf(d_2)
        val = term1 + term2
    else:
        term2 = r * K * np.exp(-r * T) * norm.cdf(-d_2)
        val = term1 + term2
    return val / 365

def rho(S, K, T, r, sigma, option_type):
    d_2 = d2(S, K, T, r, sigma)
    if option_type == 'call':
        return K * T * np.exp(-r * T) * norm.cdf(d_2) * 0.01
    else:
        return -K * T * np.exp(-r * T) * norm.cdf(-d_2) * 0.01

print(f"\n--- Table 1: Greeks (Historical Vol = {annualized_volatility:.4f}) ---")
hist_greeks_results = []
for T_days in maturities_days:
    T_years = T_days / 365.0
    for K in strikes:
        hist_greeks_results.append({
            'Days': T_days,
            'Strike': round(K, 2),
            'C_Delta': round(delta(S, K, T_years, r, annualized_volatility, 'call'), 4),
            'P_Delta': round(delta(S, K, T_years, r, annualized_volatility, 'put'), 4),
            'Gamma': round(gamma(S, K, T_years, r, annualized_volatility), 4),
            'Vega': round(vega(S, K, T_years, r, annualized_volatility), 4),
            'C_Theta': round(theta(S, K, T_years, r, annualized_volatility, 'call'), 4),
            'P_Theta': round(theta(S, K, T_years, r, annualized_volatility, 'put'), 4),
            'C_Rho': round(rho(S, K, T_years, r, annualized_volatility, 'call'), 4),
            'P_Rho': round(rho(S, K, T_years, r, annualized_volatility, 'put'), 4)
        })
hist_greeks_df = pd.DataFrame(hist_greeks_results)
print(hist_greeks_df.to_string(index=False))

hist_greeks_df.to_excel(writer, sheet_name='Greeks - Historical Vol', index=False)


print("\nFetching Live Option Chain from NSE...")

market_iv_dict = {} 
use_market_iv = False

try:
    oc_data = option_chain('ICICIBANK')
    S_live = oc_data['records']['underlyingValue']
    expiry_dates = oc_data['records']['expiryDates']
    print(f"Live Stock Price: {S_live}")
    print(f"Available Expiries: {expiry_dates[:3]}")
    
    for target_days in maturities_days:
        print(f"\n--- Processing {target_days}-day maturity ---")
        
        best_expiry = None
        min_day_diff = float('inf')
        
        for exp_date_str in expiry_dates:
            exp_date_obj = datetime.strptime(exp_date_str, '%d-%b-%Y')
            days_to_expiry = (exp_date_obj - datetime.now()).days
            
            if days_to_expiry > 0:
                day_diff = abs(days_to_expiry - target_days)
                if day_diff < min_day_diff:
                    min_day_diff = day_diff
                    best_expiry = exp_date_str
                    best_days = days_to_expiry
        
        if best_expiry is None:
            print(f"No valid expiry found for {target_days} days")
            continue
            
        print(f"Selected expiry: {best_expiry} ({best_days} days actual)")
        T_market = best_days / 365.0
        
        for K in strikes:
            print(f"  Finding IV for Strike {K:.2f}...", end=" ")
            
            closest_market_strike = None
            min_strike_diff = float('inf')
            closest_option_data = None
            
            for option in oc_data['records']['data']:
                if option['expiryDate'] == best_expiry and 'CE' in option:
                    market_strike = option['strikePrice']
                    strike_diff = abs(market_strike - K)
                    
                    if strike_diff < min_strike_diff:
                        min_strike_diff = strike_diff
                        closest_market_strike = market_strike
                        closest_option_data = option['CE']
            
            if closest_option_data and closest_option_data['lastPrice'] > 0 and min_strike_diff < K * 0.02:
                try:
                    market_price = closest_option_data['lastPrice']
                    
                    iv_result = fsolve(iv_solver, 0.3, 
                                      args=(S_live, closest_market_strike, T_market, 
                                            r, market_price, 'call'))
                    iv_calculated = iv_result[0]
                    
                    if 0.05 < iv_calculated < 2.0:
                        market_iv_dict[(K, target_days)] = iv_calculated
                        print(f"IV = {iv_calculated:.4f} (using market strike {closest_market_strike})")
                        use_market_iv = True
                    else:
                        print(f"IV out of range ({iv_calculated:.4f}), using historical vol")
                        market_iv_dict[(K, target_days)] = annualized_volatility
                except Exception as e:
                    print(f"Calculation failed: {e}, using historical vol")
                    market_iv_dict[(K, target_days)] = annualized_volatility
            else:
                print(f"No valid market data, using historical vol")
                market_iv_dict[(K, target_days)] = annualized_volatility
    
    if use_market_iv:
        print(f"\n✓ Successfully calculated market IVs for {len(market_iv_dict)} combinations")
    else:
        print("\n✗ Could not calculate market IVs, using historical volatility")
        for T_days in maturities_days:
            for K in strikes:
                market_iv_dict[(K, T_days)] = annualized_volatility

except Exception as e:
    print(f"\nError fetching NSE data: {e}")
    print("Using historical volatility for all options")
    for T_days in maturities_days:
        for K in strikes:
            market_iv_dict[(K, T_days)] = annualized_volatility

print(f"\n--- Table 2: Greeks (Using Strike-Specific Implied Vol) ---")
iv_greeks_results = []

for T_days in maturities_days:
    T_years = T_days / 365.0
    
    for K in strikes:
        strike_specific_iv = market_iv_dict.get((K, T_days), annualized_volatility)
        
        iv_greeks_results.append({
            'Days': T_days,
            'Strike': round(K, 2),
            'Implied Vol': round(strike_specific_iv, 4),  
            'C_Delta': round(delta(S, K, T_years, r, strike_specific_iv, 'call'), 4),
            'P_Delta': round(delta(S, K, T_years, r, strike_specific_iv, 'put'), 4),
            'Gamma': round(gamma(S, K, T_years, r, strike_specific_iv), 4),
            'Vega': round(vega(S, K, T_years, r, strike_specific_iv), 4),
            'C_Theta': round(theta(S, K, T_years, r, strike_specific_iv, 'call'), 4),
            'P_Theta': round(theta(S, K, T_years, r, strike_specific_iv, 'put'), 4),
            'C_Rho': round(rho(S, K, T_years, r, strike_specific_iv, 'call'), 4),
            'P_Rho': round(rho(S, K, T_years, r, strike_specific_iv, 'put'), 4)
        })

iv_greeks_df = pd.DataFrame(iv_greeks_results)
print(iv_greeks_df.to_string(index=False))

iv_greeks_df.to_excel(writer, sheet_name='Greeks - Implied Vol', index=False)

print("\n--- Greeks Comparison: Historical Vol vs Implied Vol ---")
comparison_data = []

for T_days in maturities_days:
    T_years = T_days / 365.0
    
    for K in strikes:
        hist_vol = annualized_volatility
        impl_vol = market_iv_dict.get((K, T_days), annualized_volatility)
        
        hist_call_delta = delta(S, K, T_years, r, hist_vol, 'call')
        impl_call_delta = delta(S, K, T_years, r, impl_vol, 'call')
        
        hist_gamma = gamma(S, K, T_years, r, hist_vol)
        impl_gamma = gamma(S, K, T_years, r, impl_vol)
        
        comparison_data.append({
            'Days': T_days,
            'Strike': round(K, 2),
            'Hist Vol': round(hist_vol, 4),
            'Impl Vol': round(impl_vol, 4),
            'Vol Diff %': round((impl_vol - hist_vol) / hist_vol * 100, 2),
            'Call Delta (Hist)': round(hist_call_delta, 4),
            'Call Delta (Impl)': round(impl_call_delta, 4),
            'Gamma (Hist)': round(hist_gamma, 4),
            'Gamma (Impl)': round(impl_gamma, 4)
        })

comparison_df = pd.DataFrame(comparison_data)
print(comparison_df.to_string(index=False))
comparison_df.to_excel(writer, sheet_name='Greeks Comparison', index=False)

# --- Volatility Surface (Keep mostly same but add better error handling) ---
print("\nGenerating Volatility Surface...")
all_iv_points = []

try:
    for exp_date_str in expiry_dates[:6]:  
        exp_date_obj = datetime.strptime(exp_date_str, '%d-%b-%Y')
        days_to_expiry = (exp_date_obj - datetime.now()).days
        T_market = days_to_expiry / 365.0
        
        if T_market <= 0.01: 
            continue
        
        for option in oc_data['records']['data']:
            if option['expiryDate'] == exp_date_str and 'CE' in option:
                ce = option['CE']
                
                if (ce['lastPrice'] > 0 and 
                    ce.get('totalTradedVolume', 0) > 0 and
                    0.85 * S_live <= ce['strikePrice'] <= 1.15 * S_live):
                    
                    try:
                        iv = fsolve(iv_solver, 0.3, 
                                   args=(S_live, ce['strikePrice'], T_market, 
                                         r, ce['lastPrice'], 'call'))[0]
                        
                        if 0.05 < iv < 1.5:
                            all_iv_points.append([ce['strikePrice'], days_to_expiry, iv])
                    except:
                        pass
    
    print(f"Collected {len(all_iv_points)} IV data points for surface")
    
    if len(all_iv_points) >= 10:
        iv_df = pd.DataFrame(all_iv_points, columns=['Strike', 'Days', 'IV'])
        
        iv_df.to_excel(writer, sheet_name='IV Surface Data', index=False)
        
        fig = plt.figure(figsize=(12, 8))
        ax = fig.add_subplot(111, projection='3d')
        
        surf = ax.plot_trisurf(iv_df['Strike'], iv_df['Days'], iv_df['IV'], 
                               cmap=cm.viridis, linewidth=0.1, alpha=0.8)
        
        ax.set_title('Implied Volatility Surface (ICICIBANK)', fontsize=14, fontweight='bold')
        ax.set_xlabel('Strike Price', fontsize=11)
        ax.set_ylabel('Days to Maturity', fontsize=11)
        ax.set_zlabel('Implied Volatility', fontsize=11)
        
        fig.colorbar(surf, shrink=0.5, aspect=5)
        plt.savefig('iv_surface_plot.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        print("✓ Volatility surface plot created successfully")
    else:
        print(f"✗ Only {len(all_iv_points)} data points - insufficient for surface plot")
        print("  (Need at least 10 points with liquid options)")

except Exception as e:
    print(f"Error creating volatility surface: {e}")

print("\n" + "="*50)
print("PART C: GREEKS & VOLATILITY - COMPLETE")
print("="*50)

print("\n" + "="*50)
print("PART D: PORTFOLIO & HEDGING")
print("="*50)

# Define Portfolio with 6 Options
# 1. Long 1 ATM Call (30 days)
# 2. Long 1 ATM Put (30 days)
# 3. Short 1 OTM Call at ATM+5% (30 days)
# 4. Long 1 ITM Call at ATM-5% (60 days)
# 5. Short 1 OTM Put at ATM-5% (60 days)
# 6. Long 1 ATM Call (90 days)

portfolio_positions = [
    {'type': 'call', 'strike': K_atm, 'maturity': 30, 'quantity': 1, 'description': 'Long ATM Call 30D'},
    {'type': 'put', 'strike': K_atm, 'maturity': 30, 'quantity': 1, 'description': 'Long ATM Put 30D'},
    {'type': 'call', 'strike': K_atm * 1.05, 'maturity': 30, 'quantity': -1, 'description': 'Short OTM Call 30D'},
    {'type': 'call', 'strike': K_atm * 0.95, 'maturity': 60, 'quantity': 1, 'description': 'Long ITM Call 60D'},
    {'type': 'put', 'strike': K_atm * 0.95, 'maturity': 60, 'quantity': -1, 'description': 'Short OTM Put 60D'},
    {'type': 'call', 'strike': K_atm, 'maturity': 90, 'quantity': 1, 'description': 'Long ATM Call 90D'}
]

print("\n--- Portfolio Composition (6 Options) ---")
portfolio_composition = []
for i, pos in enumerate(portfolio_positions, 1):
    T_years = pos['maturity'] / 365.0
    option_price = bsm_pricer(S, pos['strike'], T_years, r, annualized_volatility, pos['type'])
    print(f"{i}. {pos['description']}: Strike={pos['strike']:.2f}, Price={option_price:.2f}, Qty={pos['quantity']}")
    portfolio_composition.append({
        'Option #': i,
        'Description': pos['description'],
        'Type': pos['type'].capitalize(),
        'Strike': round(pos['strike'], 2),
        'Maturity (Days)': pos['maturity'],
        'Quantity': pos['quantity'],
        'Unit Price': round(option_price, 2),
        'Total Value': round(option_price * pos['quantity'], 2)
    })

portfolio_comp_df = pd.DataFrame(portfolio_composition)

port_delta = 0
port_gamma = 0
port_vega = 0
port_theta = 0
port_rho = 0

for pos in portfolio_positions:
    T_years = pos['maturity'] / 365.0
    qty = pos['quantity']
    
    port_delta += delta(S, pos['strike'], T_years, r, annualized_volatility, pos['type']) * qty
    port_gamma += gamma(S, pos['strike'], T_years, r, annualized_volatility) * qty
    port_vega += vega(S, pos['strike'], T_years, r, annualized_volatility) * qty
    port_theta += theta(S, pos['strike'], T_years, r, annualized_volatility, pos['type']) * qty
    port_rho += rho(S, pos['strike'], T_years, r, annualized_volatility, pos['type']) * qty

print(f"\n--- Portfolio Greeks (Before Hedging) ---")
print(f"Portfolio Delta: {port_delta:.4f}")
print(f"Portfolio Gamma: {port_gamma:.4f}")
print(f"Portfolio Vega:  {port_vega:.4f}")
print(f"Portfolio Theta: {port_theta:.4f}")
print(f"Portfolio Rho:   {port_rho:.4f}")

shares_hedge = -port_delta 

K_otm = K_atm * 1.05
T_hedge = 30 / 365.0
gamma_otm = gamma(S, K_otm, T_hedge, r, annualized_volatility)
options_hedge = -port_gamma / gamma_otm 

print("\n--- Hedging Strategy ---")
print(f"Delta Hedge: Trade {shares_hedge:.4f} shares of underlying stock")
print(f"Gamma Hedge: Trade {options_hedge:.4f} OTM Calls (Strike {K_otm:.2f}, 30 Days)")

hedge_opt_delta = delta(S, K_otm, T_hedge, r, annualized_volatility, 'call')
hedge_opt_gamma = gamma(S, K_otm, T_hedge, r, annualized_volatility)
hedge_opt_vega = vega(S, K_otm, T_hedge, r, annualized_volatility)
hedge_opt_theta = theta(S, K_otm, T_hedge, r, annualized_volatility, 'call')
hedge_opt_rho = rho(S, K_otm, T_hedge, r, annualized_volatility, 'call')

final_delta = port_delta + shares_hedge * 1 + (options_hedge * hedge_opt_delta)
final_gamma = port_gamma + (options_hedge * hedge_opt_gamma)
final_vega = port_vega + (options_hedge * hedge_opt_vega)
final_theta = port_theta + (options_hedge * hedge_opt_theta)
final_rho = port_rho + (options_hedge * hedge_opt_rho)

print(f"\n--- Portfolio Greeks (After Hedging) ---")
print(f"Final Hedged Delta: {final_delta:.4f} (Target: ~0)")
print(f"Final Hedged Gamma: {final_gamma:.4f} (Target: ~0)")
print(f"Final Hedged Vega:  {final_vega:.4f}")
print(f"Final Hedged Theta: {final_theta:.4f}")
print(f"Final Hedged Rho:   {final_rho:.4f}")

portfolio_comp_df.to_excel(writer, sheet_name='Portfolio Composition', index=False)

portfolio_greeks = pd.DataFrame({
    'Greek': ['Delta', 'Gamma', 'Vega', 'Theta', 'Rho'],
    'Before Hedging': [port_delta, port_gamma, port_vega, port_theta, port_rho],
    'After Hedging': [final_delta, final_gamma, final_vega, final_theta, final_rho]
})
portfolio_greeks.to_excel(writer, sheet_name='Portfolio Greeks', index=False)

hedging_positions = pd.DataFrame({
    'Instrument': ['Underlying Stock', f'OTM Call (K={K_otm:.2f}, 30D)'],
    'Quantity': [shares_hedge, options_hedge],
    'Purpose': ['Delta Hedge', 'Gamma Hedge']
})
hedging_positions.to_excel(writer, sheet_name='Hedging Positions', index=False)

print("\n--- PnL Simulation (Task 12) ---")
print("Change | Unhedged PnL | Hedged PnL")
pnl_results = []
for chg in [-0.02, -0.01, 0.01, 0.02]:
    dS = S * chg
    pnl_unhedged = (port_delta * dS) + (0.5 * port_gamma * dS**2)
    pnl_hedged = (final_delta * dS) + (0.5 * final_gamma * dS**2)
    print(f"{chg*100:>+3.0f}%   | {pnl_unhedged:>10.2f}   | {pnl_hedged:>10.2f}")
    pnl_results.append({
        'Price Change (%)': f"{chg*100:+.0f}%",
        'Unhedged PnL': round(pnl_unhedged, 2),
        'Hedged PnL': round(pnl_hedged, 2)
    })

pnl_df = pd.DataFrame(pnl_results)
pnl_df.to_excel(writer, sheet_name='PnL Simulation', index=False)

# PART E: RISK (VaR)
print("\n" + "="*50)
print("PART E: RISK (VaR CALCULATION)")
print("="*50)

ret_mean = log_returns.mean()
ret_std = log_returns.std()

var_95_param = norm.ppf(0.05, loc=ret_mean, scale=ret_std)
var_99_param = norm.ppf(0.01, loc=ret_mean, scale=ret_std)

var_95_hist = np.percentile(log_returns, 5)
var_99_hist = np.percentile(log_returns, 1)

print("\n--- Stock Returns VaR (1-day) ---")
print("\nParametric VaR:")
print(f"95% Conf: {var_95_param:.4f} ({var_95_param*100:.2f}%)")
print(f"99% Conf: {var_99_param:.4f} ({var_99_param*100:.2f}%)")

print("\nHistorical VaR:")
print(f"95% Conf: {var_95_hist:.4f} ({var_95_hist*100:.2f}%)")
print(f"99% Conf: {var_99_hist:.4f} ({var_99_hist*100:.2f}%)")

print("\n" + "="*50)
print("PORTFOLIO VaR COMPARISON (Hedged vs Unhedged)")
print("="*50)

unhedged_portfolio_value = 0
for pos in portfolio_positions:
    T_years = pos['maturity'] / 365.0
    option_price = bsm_pricer(S, pos['strike'], T_years, r, annualized_volatility, pos['type'])
    unhedged_portfolio_value += option_price * pos['quantity']

print(f"\nUnhedged Portfolio Value: {unhedged_portfolio_value:.2f}")

if abs(unhedged_portfolio_value) < 1.0:
    print(f"\n WARNING: Portfolio value is very small ({unhedged_portfolio_value:.4f})")
    print("    Returns calculation may be unreliable. Consider adjusting portfolio positions.")

min_portfolio_value = max(abs(unhedged_portfolio_value), 1.0)  

if min_portfolio_value > abs(unhedged_portfolio_value):
    print(f"    Using normalized value of {min_portfolio_value:.2f} for return calculations.")

portfolio_returns_unhedged = []
portfolio_returns_hedged = []

hist_prices = hist_data['Close'].tail(60).values

for i in range(1, len(hist_prices)):
    S_old = hist_prices[i-1]
    S_new = hist_prices[i]
    price_change = S_new - S_old
    
    pnl_unhedged = (port_delta * price_change) + (0.5 * port_gamma * price_change**2)
    ret_unhedged = pnl_unhedged / min_portfolio_value
    portfolio_returns_unhedged.append(ret_unhedged)
    
    pnl_hedged = (final_delta * price_change) + (0.5 * final_gamma * price_change**2)
    ret_hedged = pnl_hedged / min_portfolio_value 
    portfolio_returns_hedged.append(ret_hedged)

portfolio_returns_unhedged = np.array(portfolio_returns_unhedged)
portfolio_returns_hedged = np.array(portfolio_returns_hedged)

unhedged_mean = portfolio_returns_unhedged.mean()
unhedged_std = portfolio_returns_unhedged.std()
var_95_param_unhedged = norm.ppf(0.05, loc=unhedged_mean, scale=unhedged_std)
var_99_param_unhedged = norm.ppf(0.01, loc=unhedged_mean, scale=unhedged_std)

var_95_hist_unhedged = np.percentile(portfolio_returns_unhedged, 5)
var_99_hist_unhedged = np.percentile(portfolio_returns_unhedged, 1)

hedged_mean = portfolio_returns_hedged.mean()
hedged_std = portfolio_returns_hedged.std()
var_95_param_hedged = norm.ppf(0.05, loc=hedged_mean, scale=hedged_std)
var_99_param_hedged = norm.ppf(0.01, loc=hedged_mean, scale=hedged_std)

var_95_hist_hedged = np.percentile(portfolio_returns_hedged, 5)
var_99_hist_hedged = np.percentile(portfolio_returns_hedged, 1)

print("\n--- UNHEDGED PORTFOLIO VaR (1-day) ---")
print("\nParametric VaR:")
print(f"95% Conf: {var_95_param_unhedged:.4f} ({var_95_param_unhedged*100:.2f}%)")
print(f"99% Conf: {var_99_param_unhedged:.4f} ({var_99_param_unhedged*100:.2f}%)")
print("\nHistorical VaR:")
print(f"95% Conf: {var_95_hist_unhedged:.4f} ({var_95_hist_unhedged*100:.2f}%)")
print(f"99% Conf: {var_99_hist_unhedged:.4f} ({var_99_hist_unhedged*100:.2f}%)")

print("\n--- HEDGED PORTFOLIO VaR (1-day) ---")
print("\nParametric VaR:")
print(f"95% Conf: {var_95_param_hedged:.4f} ({var_95_param_hedged*100:.2f}%)")
print(f"99% Conf: {var_99_param_hedged:.4f} ({var_99_param_hedged*100:.2f}%)")
print("\nHistorical VaR:")
print(f"95% Conf: {var_95_hist_hedged:.4f} ({var_95_hist_hedged*100:.2f}%)")
print(f"99% Conf: {var_99_hist_hedged:.4f} ({var_99_hist_hedged*100:.2f}%)")

print("\n--- VaR COMPARISON (Hedged vs Unhedged) ---")
print(f"\nRisk Reduction (Parametric 95%): {((var_95_param_unhedged - var_95_param_hedged) / abs(var_95_param_unhedged) * 100):.2f}%")
print(f"Risk Reduction (Parametric 99%): {((var_99_param_unhedged - var_99_param_hedged) / abs(var_99_param_unhedged) * 100):.2f}%")
print(f"Risk Reduction (Historical 95%): {((var_95_hist_unhedged - var_95_hist_hedged) / abs(var_95_hist_unhedged) * 100):.2f}%")
print(f"Risk Reduction (Historical 99%): {((var_99_hist_unhedged - var_99_hist_hedged) / abs(var_99_hist_unhedged) * 100):.2f}%")

stock_var_results = pd.DataFrame({
    'Method': ['Parametric', 'Parametric', 'Historical', 'Historical'],
    'Confidence Level': ['95%', '99%', '95%', '99%'],
    'VaR (Decimal)': [var_95_param, var_99_param, var_95_hist, var_99_hist],
    'VaR (%)': [var_95_param*100, var_99_param*100, var_95_hist*100, var_99_hist*100]
})
stock_var_results.to_excel(writer, sheet_name='Stock VaR', index=False)

portfolio_var_comparison = pd.DataFrame({
    'Method': ['Parametric', 'Parametric', 'Historical', 'Historical'],
    'Confidence Level': ['95%', '99%', '95%', '99%'],
    'Unhedged VaR (%)': [var_95_param_unhedged*100, var_99_param_unhedged*100, 
                          var_95_hist_unhedged*100, var_99_hist_unhedged*100],
    'Hedged VaR (%)': [var_95_param_hedged*100, var_99_param_hedged*100, 
                        var_95_hist_hedged*100, var_99_hist_hedged*100],
    'Risk Reduction (%)': [
        (var_95_param_unhedged - var_95_param_hedged) / abs(var_95_param_unhedged) * 100,
        (var_99_param_unhedged - var_99_param_hedged) / abs(var_99_param_unhedged) * 100,
        (var_95_hist_unhedged - var_95_hist_hedged) / abs(var_95_hist_unhedged) * 100,
        (var_99_hist_unhedged - var_99_hist_hedged) / abs(var_99_hist_unhedged) * 100
    ]
})
portfolio_var_comparison.to_excel(writer, sheet_name='Portfolio VaR Comparison', index=False)

print("\nAssignment Tasks Complete.")

writer.close()
print(f"\nAll tables exported to: {excel_filename}")
print("All plots saved as PNG files in the current directory.")