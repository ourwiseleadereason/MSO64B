import pandas as pd
with open("1.9/Frequency.csv", 'r') as file:
    lines = file.readlines()
lines = [line.strip() for line in lines]
cleaned_data = [entry.rstrip(', ').rstrip() for entry in lines]
for i in range(len(cleaned_data)):
    cleaned_data[i]+='\n'

with open("bruh_moment.csv", 'w') as file:
    file.writelines(cleaned_data)

def process(x):
    entries=x.split(' ')
    multiplier=1
    if entries[1]=="kHz":
        multiplier=float(1/1000)
    elif entries[1]=="ns":
        multiplier=1e9
    elif entries[1]=="us":
        multiplier==1e6
    return float(entries[0])/multiplier

def generate_info(df, scl, sda):
    return f"{df['Measurement']}({sda},{scl})"

df=pd.read_csv("bruh_moment.csv")
df2=pd.DataFrame()
df2['Value']=df["Mean'"].apply(process)
df2['Mean']=df["Mean'"].apply(process)
df2['Min']=df["Min'"].apply(process)
df2['Max']=df["Max'"].apply(process)
df2['St Dev']=df["Std Dev'"].apply(process)
df2['Count']=df["Population'"]
sda="CH2"
scl="CH1"
df2['info'] = df['Measurement'] + f"({sda},{scl})"
df2.to_csv('output.csv', index=False)
