import psycopg2
import xlsxwriter
import openpyxl

hostname = 'localhost'
db = 'python-sql-assignment'
username = 'postgres'
pwd = '17JE003089'
port_id = 5432


class Assignment:
    def __init__(self):
        self.conn = psycopg2.connect(
            user=username,
            host=hostname,
            database=db,
            password=pwd
        )
        self.workbook = xlsxwriter.Workbook("sqlite-assignment.xlsx")

    def __del__(self):
        self.workbook.close()
        self.conn.close()

    def list_emp_details(self):
        query1 = "select empno, ename, mgr from emp"
        cur = self.conn.cursor()
        cur.execute(query1)
        rows = cur.fetchall()

        worksheet1 = self.workbook.add_worksheet("q1")

        rowIdx = 2

        worksheet1.write('A1', 'employee number')
        worksheet1.write('B1', 'names')
        worksheet1.write('C1', 'managers')

        for r in rows:
            worksheet1.write('A' + str(rowIdx), r[0])
            worksheet1.write('B' + str(rowIdx), r[1])
            worksheet1.write('C' + str(rowIdx), r[2])
            rowIdx += 1
            print(r)

        cur.close()

    def list_total_compensation(self):
        query1 = '''select
        	emp.ename,
        	emp.empno,
        	dept.dname,
         	emp.sal,
        	case
        		when jobhist.enddate is null then extract(year from age(CURRENT_TIMESTAMP, jobhist.startdate)) * 12 +
        		extract(month from age(CURRENT_TIMESTAMP, jobhist.startdate))
        		else extract(year from age(jobhist.enddate, jobhist.startdate)) * 12 +
        		extract(month from age(jobhist.enddate, jobhist.startdate))
        	end
        from emp
        join dept on emp.deptno = dept.deptno
        join jobhist on emp.empno = jobhist.empno;'''

        cur1 = self.conn.cursor()
        cur1.execute(query1)
        rows1 = cur1.fetchall()

        worksheet2 = self.workbook.add_worksheet("q2")

        worksheet2.write('A1', 'Emp Name')
        worksheet2.write('B1', 'Emp No')
        worksheet2.write('C1', 'Dept Name')
        worksheet2.write('D1', 'Total Compensation')
        worksheet2.write('E1', 'Month Spent in Organisation')

        rowIdx = 2

        for i in rows1:
            worksheet2.write('A' + str(rowIdx), i[0])
            worksheet2.write('B' + str(rowIdx), i[1])
            worksheet2.write('C' + str(rowIdx), i[2])
            worksheet2.write('D' + str(rowIdx), i[3])
            worksheet2.write('E' + str(rowIdx), i[4])
            rowIdx += 1
            print(i)

        cur1.close()

    def read_db(self):
        wb = openpyxl.load_workbook("sql-assignment.xlsx")
        sh1 = wb['q2']

        row = sh1.max_row
        col = sh1.max_column

        curr2 = self.conn.cursor()

        curr2.execute("insert into postgresDB (ename, empno, dname, sal, month_diff) values (%s, %s, %s, %s, %s)",
                      ("SMITH", 7369, "RESEARCH", 800, 494))

        for i in range(2, row + 1):
            li = []
            for j in range(1, col + 1):
                li.append(sh1.cell(i, j).value)
            print(f"#{i}")
            print(li)
            curr2.execute("insert into postgresDB (ename, empno, dname, sal, month_diff) values (%s, %s, %s, %s, %s)",
                          (li[0], li[1], li[2], li[3], li[4]))

        self.conn.commit()
        curr2.close()
        curx = self.conn.cursor()
        curx.execute("select * from postgresDB")
        rows = curx.fetchall()
        print(rows)

        curx.close()

    def list_total_compensation_atdept(self):
        cur3 = self.conn.cursor()

        query3 = '''select
        	deptno,
        	dept.dname,
        	sum(sal)
        from postgresDB
        join dept on postgresDB.dname = dept.dname
        group by deptno;'''

        cur3.execute(query3)
        rows2 = cur3.fetchall()

        worksheet3 = self.workbook.add_worksheet("q4")

        worksheet3.write('A1', 'Dept No')
        worksheet3.write('B1', 'Dept Name')
        worksheet3.write('C1', 'Total Compensation')

        rowIdx = 2

        for r in rows2:
            worksheet3.write('A' + str(rowIdx), r[0])
            worksheet3.write('B' + str(rowIdx), r[1])
            worksheet3.write('C' + str(rowIdx), r[2])
            rowIdx += 1
            print(r)

        cur3.close()


assignment = Assignment()
assignment.list_emp_details()
