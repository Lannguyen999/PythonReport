import os
from dotenv import load_dotenv
import mysql.connector
from mysql.connector import Error

load_dotenv()

def mysql_query(sql):
  try:
    connection = mysql.connector.connect(
      host = os.getenv('MYSQL_HOST'),
      user = os.getenv('MYSQL_USER'),
      password = os.getenv('MYSQL_PASSWORD')
    )
    if connection.is_connected():
      cursor = connection.cursor()
      cursor.execute(sql)
      result = cursor.fetchall()
      return result
  except Error as e:
    print("Error when connect to MySQL", e)
  finally:
    if connection.is_connected():
      cursor.close()
      connection.close()
      print("Connection closed")
