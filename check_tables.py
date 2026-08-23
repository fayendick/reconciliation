import sqlite3
c = sqlite3.connect(r'C:\Users\ndick.faye\Documents\RECONCOR\reconciliation\data\base.db')
tables = [r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")]
for t in tables:
    print(t)
