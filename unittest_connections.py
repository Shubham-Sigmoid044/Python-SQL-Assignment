import pytest
import psycopg2

hostname = 'localhost'
db = 'python-sql-assignment'
username = 'postgres'
pwd = '17JE003089'
port_id = 5432

conn = None


def setup_module(module):
    global conn

    conn = psycopg2.connect(
        user=username,
        host=hostname,
        database=db,
        password=pwd
    )


def teardown_module(module):
    conn.close()


def test_emp_details():
    query1 = "select empno, ename, mgr from emp"  # SQL query to list all employee details
    cur = conn.cursor()
    cur.execute(query1)
    rows = cur.fetchone()  # to execute the query using given cursor

    cur.close()
    print(rows)
    assert rows[0] == 7369
    assert rows[1] == 'SMITH'
    assert rows[2] == 7902


def test_list_total_compensation():
    # Created the query using CASE statement in which if there hiring date is NULL then we are subtracting with
    # current date else subtracting with leaving date.
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

    cur1 = conn.cursor()
    cur1.execute(query1)
    rows1 = cur1.fetchone()  # to execute the query using given cursor

    cur1.close()

    print(rows1)
    assert rows1[0] == 'SMITH'
    assert rows1[1] == 7369
    assert rows1[2] == 'RESEARCH'
    assert rows1[3] == 800


def test_list_total_compensation_atdept():
    cur3 = conn.cursor()

    # Group by dept number to calculate total compensation per dept
    query3 = '''select
            	deptno,
            	dept.dname,
            	sum(sal)
            from postgresDB
            join dept on postgresDB.dname = dept.dname
            group by deptno;'''

    cur3.execute(query3)
    rows2 = cur3.fetchone()  # to execute the query using given cursor

    print(rows2)

    assert rows2[0] == 10
    assert rows2[1] == 'ACCOUNTING'
    assert rows2[2] == 8750
