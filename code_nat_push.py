import openpyxl

# ambil file excel
file_path = 'push.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Initialize the NAT config
nat_config = []

# melakukan perulanagn pada setiap baris excel
for row in sheet.iter_rows(min_row=2, values_only=True):
    location = row[1]
    inside_ip = row[3]
    outside_ip = row[4]
    
    # pada lokasi ganti spasi jadi underscore
    name = location.replace(' ', '_')
    
    # Create NAT config cisco nya
    nat_entry = f"ip nat name {name} inside source static {inside_ip} {outside_ip}"
    nat_config.append(nat_entry)

# Save NAT config ke dalam text file
config_file_path = 'nat_config.txt'
with open(config_file_path, 'w') as config_file:
    config_file.write('\n'.join(nat_config))

print("NAT configuration has been generated and saved. Cheers made by -__SgT__-")
