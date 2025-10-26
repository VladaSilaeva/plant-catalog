import streamlit as st
import sqlite3

st.title("My plant-catalog app")


@st.cache_resource
def get_connection():
    return sqlite3.connect("plants.db", check_same_thread=False)

connection = get_connection()
cursor = connection.cursor()


cursor.execute('''
CREATE TABLE IF NOT EXISTS Plants (
id INTEGER PRIMARY KEY,
name TEXT NOT NULL,
rus_name TEXT
)
''')

connection.commit()



if st.button("Add row"):
    cursor.execute('INSERT INTO Plants (name) VALUES (?)', ('Oxalis',))
    connection.commit()

if st.button("Update row"):
    cursor.execute('UPDATE Plants SET rus_name = ? WHERE name = ?', ('Кислица', 'Oxalis',))
    connection.commit()

cursor.execute("SELECT * FROM Plants")
rows = cursor.fetchall()
st.write(rows)
