import pandas as pd

# Read the excel file
df_original = pd.read_excel('data.xlsx')


# Function to process the data with given h2_prod_limits and h2_cons_limits
def process_data(h2_prod_limits, h2_cons_limits, df, gains_df):
    # Replace the hardcoded limits with the given limits
    h2_prod_upper_limit, h2_prod_lower_limit = h2_prod_limits
    h2_cons_upper_limit, h2_cons_lower_limit = h2_cons_limits
    

        
    # initialize state of storage to 0
    state_of_storage = 0

    # iterate over the rows of the dataframe
    for index, row in df.iterrows():
        # check if state of storage is less than 64.1 (1 cylinder)
        if state_of_storage < 64.1:
            # calculate H2 production
            h2_prod = row['surplus'] * 0.25 * 10 / 50 / 1000 * 0.98
            # check if H2 production is less than 0.125 Nm3/15min
            if h2_prod < h2_prod_lower_limit:
                h2_prod = 0
            if h2_prod > h2_prod_upper_limit:
                h2_prod = h2_prod_upper_limit
            # check if state of storage plus H2 production is less than or equal to 64.1
            if state_of_storage + h2_prod <= 64.1:
                # update state of storage
                state_of_storage += h2_prod
            else:
                # set H2 production to the remaining capacity in the storage tank
                h2_prod = 64.1 - state_of_storage
                # update state of storage
                state_of_storage = 64.1

        else:
            # set H2 production to 0 if state of storage is at maximum capacity
            h2_prod = 0

        # check if state of storage is greater than 0
        if state_of_storage > 0:
            # calculate H2 consumption
            h2_cons = row['insufficiency'] / 1000 /0.98 * 0.25 * 10.3 / 14.2 
            # check if H2 consumption is less than 0.0644 Nm3/15min (5% of 5.15Nm3/h)
            if h2_cons < h2_cons_lower_limit:
                h2_cons = 0
            if h2_cons > h2_cons_upper_limit:
                h2_cons = h2_cons_upper_limit
            # check if state of storage - H2 consumption is greater than or equal to 0
            if state_of_storage - h2_cons >= 0:
                # update state of storage
                state_of_storage -= h2_cons
            else:
                # set H2 consumption to the remaining state of storage
                h2_cons = state_of_storage
                # update state of storage
                state_of_storage = 0

        else:
            # set H2 consumption to 0 if state of storage is at minimum capacity
            h2_cons = 0

        # calculate grid in kW 
        grid = row['insufficiency'] /1000 - h2_cons  / 10.3 * 14.2 * 0.98 /0.25

        # update the 'H2 production', 'H2 consumption', 'State of storage' and 'Grid kW' columns
        df.at[index, 'H2 production'] = h2_prod
        df.at[index, 'H2 consumption'] = h2_cons
        df.at[index, 'State of storage'] = state_of_storage
        




    # Set the initial value of soc (state of charge) to 50 kWh 
    soc_prev = 50

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():

        # Calculate the Charge column, battery charge in kWh per 15min
        charge = (row['surplus'] / 1000 * 0.98 * 0.25) - (row['H2 production'] * 50 / 10 * 0.25)  

        # Calculate the Discharge column, battery discharge in kWh, this discharge also include grid 
        discharge = ((row['insufficiency'] / 1000 * 0.25) - (row['H2 consumption'] * 14.2 / 10.3 * 0.25 * 0.98))/0.98

        # Calculate the Soc column, Soc in kWh 
        soc = soc_prev + charge - discharge

        # Check if the calculated Soc value falls outside the range of 5 to 50
        if soc < 5:
            soc = 5
        elif soc > 50:
            soc = 50

        # Calculate the RealDischarge column, RealDischarge in kWh, will remove grid 
        realdischarge = soc_prev - soc

        # Check if the calculated RealDischarge value is less than 0
        if realdischarge < 0:
            realdischarge = 0

        # Calculate the Grid column, Grid in kW 
        #grid2 = (row['insufficiency'] / 1000) - (row['H2 consumption'] * 14.2 / 10.3 ) - realdischarge
        grid2 = discharge - realdischarge

        # Check if the calculated Grid value is less than 0
        if grid2 < 0:
            grid2 = 0

        # Update the Soc_prev variable for the next iteration
        soc_prev = soc

        # Update the DataFrame with the new values
        df.at[index, 'Charge'] = charge
        df.at[index, 'Discharge'] = discharge
        df.at[index, 'SOC'] = soc
        df.at[index, 'RealDischarge'] = realdischarge
        df.at[index, 'Grid kW 2'] = grid2



    # Calculate the electricity gain
    electricity_gain = (df['H2 consumption'].sum() * 14.2 / 10.3) + df['RealDischarge'].sum()

    #add self sufficient rate
    self_sufficient_rate = ((df['H2 consumption'].sum() * 14.2 / 10.3) + df['RealDischarge'].sum())/((df['grid2'].sum()*0.25)+(df['H2 consumption'].sum() * 14.2 / 10.3) + df['RealDischarge'].sum())
    
    # Append the results to the gains DataFrame
    gains_df.loc[len(gains_df)] = [h2_prod_upper_limit, h2_cons_upper_limit, electricity_gain]





# Create a DataFrame to store the electricity gains and corresponding h2_prod and h2_cons values
gains_df = pd.DataFrame(columns=['h2_prod', 'h2_cons', 'electricity_gain', 'self_sufficient_rate')

# Iterate over the specified h2_prod and h2_cons ranges and calculate electricity gains
for h2_prod_upper_limit in [x * 0.25 for x in range(1, 11)]:
    h2_prod_lower_limit = h2_prod_upper_limit * 0.1

    for h2_cons_upper_limit in [x * 0.363 for x in range(1, 8)]:
        h2_cons_lower_limit = h2_cons_upper_limit * 0.05

        # Create a copy of the original DataFrame for processing
        df = df_original.copy()

        # Process the data and update the gains DataFrame for the current pair of h2_prod and h2_cons
        process_data((h2_prod_upper_limit, h2_prod_lower_limit),
                     (h2_cons_upper_limit, h2_cons_lower_limit), df, gains_df)

# Save the electricity gains DataFrame to an Excel file
gains_df.to_excel('electricity_gains.xlsx', index=False)