import sqlite3

conn = sqlite3.connect(r'C:\Users\DavidSant\effi-mail\effi_mail.db')
c = conn.cursor()

# Get all tables
c.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = [r[0] for r in c.fetchall()]
print("TABLES:", tables)
print()

# Show schema and row count for each table
for t in tables:
    c.execute(f"PRAGMA table_info({t})")
    cols = [r[1] for r in c.fetchall()]
    c.execute(f"SELECT COUNT(*) FROM {t}")
    count = c.fetchone()[0]
    print(f"{t} ({count} rows): {cols}")

print()

# Sample data from key tables
print("--- CLIENTS ---")
c.execute("SELECT * FROM clients")
for row in c.fetchall():
    print(row)

print()
print("--- MATTERS ---")
c.execute("SELECT * FROM matters")
for row in c.fetchall():
    print(row)

print()
print("--- DOMAINS (first 10) ---")
c.execute("SELECT * FROM domains LIMIT 10")
for row in c.fetchall():
    print(row)

conn.close()
