import openpyxl

# Use '=' instead of '-'
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "EV Market Research"

headers = ["Company Name", "Founders", "Headquarters", "Website Link", "USP", "Starting Price (INR)"]
sheet.append(headers)

# Use '=' instead of '-'
ev_data = [
    ["Ola Electric", "Bhavish Aggarwal", "Bengaluru, Karnataka", "https://olaelectric.com", "MoveOS features & charging network", 84999],
    ["Ather Energy", "Tarun Mehta & Swapnil Jain", "Bengaluru, Karnataka", "https://atherenergy.com", "Aluminium chassis & Google Maps ecosystem", 121499],
    ["TVS Motor Company", "T. V. Sundram Iyengar", "Chennai, Tamil Nadu", "https://tvsmotor.com", "Dual-helmet boot storage & high comfort", 100822],
    ["Bajaj Auto", "Jamnalal Bajaj", "Pune, Maharashtra", "https://chetak.com", "Sheet metal body architecture & retro design", 102498],
    ["Honda India", "Soichiro Honda", "Gurugram, Haryana", "https://honda2wheelersindia.com", "Swappable battery ecosystem framework", 110000]
]

for row in ev_data:
    sheet.append(row)

wb.save("EV_Scooter_Data.xlsx")
print("✔️ 'EV_Scooter_Data.xlsx' generated successfully!")

