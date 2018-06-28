import cx_Oracle

connection = cx_Oracle.connect("PANKAJ", "PANKAJ", "192.168.50.103:1521/RMS14")
queries = ["""select 1 a, 2 b, 3 c
		from dual
		union
		select 4 a, 5 b, 6 c
		from dual"""
		,
		"""select 11 first, 12 second, 13 third from dual
		union
		select 21 first, 22 second, 23 third from dual"""]
cur = connection.cursor()
for query in queries:
	cur.execute(query)
	result = cur.fetchall()
	for record in result:
		for item in record:
			print(item, end="")
		print("\r")
connection.close()


