import numpy as np
from scipy.stats import chi2_contingency

def monte_carlo_chi_square(contingency_table, num_simulations=1000):
    observed_chi2, observed_p_value, _, _ = chi2_contingency(contingency_table)
    
    chi2_values = []
    for _ in range(num_simulations):
        shuffled_table = np.random.permutation(contingency_table)
        chi2, p_value, _, _ = chi2_contingency(shuffled_table)
        chi2_values.append(chi2)
    
    # Calculate the simulated p-value
    simulated_p_value = (np.sum(chi2_values >= observed_chi2) + 1) / (num_simulations + 1)
    
    return observed_chi2, observed_p_value, simulated_p_value


