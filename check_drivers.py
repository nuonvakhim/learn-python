"""
Helper script to check available ODBC drivers on your system.
Run this to see what Access drivers are available.
"""
import pyodbc

print("Checking available ODBC drivers...\n")
drivers = pyodbc.drivers()

if not drivers:
    print("No ODBC drivers found!")
else:
    print("Available ODBC drivers:")
    print("-" * 60)
    for driver in sorted(drivers):
        print(f"  - {driver}")
    
    print("\n" + "-" * 60)
    access_drivers = [d for d in drivers if 'access' in d.lower() or 'mdb' in d.lower()]
    
    if access_drivers:
        print("\n✓ Microsoft Access drivers found:")
        for driver in access_drivers:
            print(f"  - {driver}")
    else:
        print("\n✗ No Microsoft Access drivers found!")
        print("\nYou need to install Microsoft Access Database Engine:")
        print("Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        print("\nIMPORTANT: Install the version (32-bit or 64-bit) that matches your Python.")
