import pandas as pd

# Sample DataFrame
df = pd.DataFrame({
    'Name': ['Comrade Eason', 'Very Very Long Name Indeed'],
    'Message': ['Hello', 'This is a very long message that might not fit.']
})

# Write to Excel with xlsxwriter
with pd.ExcelWriter('auto_sized.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Autofit columns based on content
    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, column_len + 2)  # Add a little extra space