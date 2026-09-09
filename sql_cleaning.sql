import pandas as pd
import sqlite3
df = pd.read_excel("/content/orders_raw.xlsx")
conn = sqlite3.connect("order_raw.db")

df.to_sql("order_raw", conn, if_exists="replace", index=False)

print(pd.read_sql_query( "SELECT COUNT(*) AS tot_no_of_rows FROM order_raw",conn))
print(pd.read_sql_query("SELECT * FROM order_raw LIMIT 10",conn))
conn.commit()

print("-------------------------------------------------------------------")
print("---------------------- Check duplicate OrderID values-------------- ")
print(pd.read_sql_query("""SELECT
OrderID,
COUNT(*) AS Duplicate_Count
FROM order_raw
GROUP BY OrderID
HAVING COUNT(*) > 1
ORDER BY Duplicate_Count DESC""",conn))


print("-------------------------------------------------------------------")
print("---------------------Check missing values -------------------------")
print(pd.read_sql_query(""" SELECT
SUM(CASE WHEN CustomerName IS NULL OR TRIM(CustomerName)='' THEN 1 ELSE 0 END)AS Missing_CustomerName,
SUM(CASE WHEN Email IS NULL OR TRIM(Email)='' THEN 1 ELSE 0 END )AS Missing_Email,
SUM(CASE WHEN Phone IS NULL OR TRIM(Phone)='' THEN 1 ELSE 0 END)AS Missing_Phone,
SUM(CASE WHEN City IS NULL OR TRIM(City)='' THEN 1 ELSE 0 END)AS Missing_City,
SUM(CASE WHEN OrderDate IS NULL OR TRIM(OrderDate)='' THEN 1 ELSE 0 END)AS Missing_OrderDate,
SUM(CASE WHEN Amount IS NULL OR TRIM(Amount)='' THEN 1 ELSE 0 END)AS Missing_Amount,
SUM(CASE WHEN Status IS NULL OR TRIM(Status)='' THEN 1 ELSE 0 END) AS Missing_Status
FROM order_raw""",conn
))
conn.commit()


print("------------------------------------------------------------------")
print("---------------------- Create a staging table----------------------")
conn.execute("DROP TABLE IF EXISTS orders_staging")
conn.execute("""CREATE TABLE orders_staging AS SELECT * FROM order_raw""")
conn.commit()
print(pd.read_sql_query("SELECT COUNT(*) AS Staging_Rows FROM orders_staging ",conn))


print("------------------------------------------------------------------")
print("----------------------Remove exact duplicate rows ---------------")
conn.execute("""DELETE FROM orders_staging WHERE rowid NOT IN(SELECT MIN(rowid)
FROM orders_staging
GROUP BY
OrderID,
CustomerName,
Email,
Phone,
City,
OrderDate,
Amount,
Status)""")
conn.commit()
print(pd.read_sql_query("""SELECT COUNT(*) AS Rows_After_Duplicate_Removal FROM orders_staging """,conn))




print("----------------------------------------------------------------")
print("----------------------Remove unwanted spaces -------------------")
conn.execute("""UPDATE orders_staging SET CustomerName=TRIM(CustomerName),Email=TRIM(Email),Phone=TRIM(Phone),
City=TRIM(City),OrderDate=TRIM(OrderDate),Status=TRIM(Status),Amount=Trim(Amount)""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 10",conn))



print("----------------------------------------------------------------")
print("-----------------Standardize Customer Names ----------------")
conn.execute("""UPDATE orders_staging SET CustomerName=UPPER(SUBSTR(CustomerName,1,1))|| LOWER(SUBSTR(CustomerName, 2, INSTR(CustomerName, ' ') - 2))||' ' || UPPER(SUBSTR(CustomerName,INSTR(CustomerName,' ')+1,1))|| LOWER(SUBSTR(CustomerName,INSTR(CustomerName,' ')+2)) WHERE CustomerNAME IS NOT NULL and INSTR(CustomerName,'')>0""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 10",conn))



print("-----------------------------------------------------------------")
print("----------------Standardize City names --------------------------")
conn.execute("""UPDATE orders_staging SET City= CASE UPPER(TRIM(City))
WHEN 'CHENNAI' THEN 'Chennai'
WHEN 'BENGALURU' THEN 'Bengaluru'
WHEN 'COIMBATORE' THEN 'Coimbatore'
WHEN 'MUMBAI' THEN 'Mumbai'
WHEN 'DELHI' THEN 'Delhi'
WHEN 'HYDERABAD' THEN 'Hyderabad'
WHEN 'PUNE' THEN 'Pune'
WHEN 'KOLKATA' THEN 'Kolkata'
ELSE City
END WHERE City IS NOT NULL""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 10",conn))


print("-----------------------------------------------------------------")
print("------------------Convert Email to lowercase --------------------")
conn.execute("""UPDATE orders_staging SET Email=LOWER(TRIM(Email)) WHERE Email IS NOT NULL""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 10",conn))



print("-----------------------------------------------------------------")
print("---------------------Standardize Status -------------------------")
conn.execute("""UPDATE orders_staging SET Status= CASE UPPER(TRIM(Status))
WHEN 'COMPLETED' THEN 'Completed'
WHEN 'CANCELLED' THEN 'Cancelled'
WHEN 'CANCELED' THEN 'Canceled'
WHEN 'PENDING' THEN 'Pending'
WHEN 'REFUNDED' THEN 'Refunded'
WHEN 'N/A' THEN 'NULL'
WHEN '' THEN 'NULL'
ELSE Status
END""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 10",conn))

print("------------------------------------------------------------------")
print(pd.read_sql_query("SELECT Status,COUNT(*)AS tot from orders_staging GROUP BY Status ORDER BY tot DESC",conn))


print("-------------------------------------------------------------------")
print("--------------------Handle missing values -------------------------")
conn.execute("""UPDATE orders_staging SET CustomerName='Unknown Customer' WHERE CustomerName IS NULL OR TRIM(CustomerName)='' """)
conn.execute("""UPDATE orders_staging SET  City='Unknown' WHERE City  IS NULL OR TRIM(City)='' """)
conn.execute("""UPDATE orders_staging SET Status='Unknown' WHERE Status IS NULL OR TRIM(Status)='' """)
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))



print("-------------------------------------------------------------------")
print("----------------------Standardize OrderDate------------------------")
conn.execute("""UPDATE orders_staging SET OrderDate=CASE
WHEN OrderDate LIKE '____-__-__' THEN OrderDate
WHEN OrderDate LIKE '__-__-____' THEN SUBSTR(OrderDate,7,4)||'-'||SUBSTR(OrderDate,1,2)||'-'||SUBSTR(OrderDate,4,2)
WHEN OrderDate LIKE '__-__-____' THEN SUBSTR(OrderDate,7,4)||'-'||SUBSTR(OrderDate,4,2)||'-'||SUBSTR(OrderDate,1,2)
ELSE NULL
END WHERE OrderDate IS NOT NULL""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("------------Validate Email ----------------------------------------")
conn.execute("ALTER TABLE orders_staging ADD COLUMN Email_Valid INTEGER")
conn.commit()
conn.execute("""UPDATE orders_staging SET Email_Valid=CASE
WHEN Email=NULL THEN 0
WHEN Email LIKE '%__@__%.__%' AND Email NOT LIKE '% %' THEN 1 ELSE 0
END """)
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("--------------------------------------------------------------------")
print("-----------------Clean Phone numbers -------------------------------")
conn.execute("UPDATE orders_staging SET Phone=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Phone,'(',''),')',''),'-',''),' ',''), '+','') WHERE Phone IS NOT NULL")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("-----------------Remove +91 country code --------------------------")
conn.execute("""UPDATE orders_staging SET Phone=SUBSTR(Phone,3) WHERE Phone IS NOT NULL AND LENGTH(Phone)=12 AND SUBSTR(Phone ,1,2)='91'""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("------------------Validate Phone numbers --------------------------")
conn.execute("ALTER TABLE orders_staging ADD COLUMN Phone_Valid  INTEGER")
conn.commit()
conn.execute("""UPDATE orders_staging SET Phone_Valid =CASE
WHEN Phone IS NOT NULL AND
LENGTH(Phone)=10 AND
Phone NOT LIKE '*[^0-9]*'
THEN 1 ELSE 0
END""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))



print("-------------------------------------------------------------------")
print("--------------------Clean Amount ----------------------------------")
conn.execute("""UPDATE orders_staging SET Amount=REPLACE(REPLACE(REPLACE(Amount, '₹', ''),',',''),' ' ,'') WHERE Amount IS NOT NULL;""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("-------------------Create numeric Amount column -------------------")
conn.execute("ALTER TABLE orders_staging ADD COLUMN Amount_Clean REAL ")
conn.execute("UPDATE orders_staging SET Amount_Clean=CAST(Amount AS REAL) WHERE Amount IS NOT NULL")
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("--------------------- Identify Amount problems --------------------")
conn.execute("ALTER TABLE orders_staging ADD COLUMN Amount_Flag  TEXT")
conn.execute("""UPDATE orders_staging SET Amount_Flag=CASE
WHEN Amount_Clean IS NULL THEN 'Missing'
WHEN Amount_Clean<0 THEN 'Negativ'
WHEN Amount_Clean >20000 THEN 'Outlier'
ELSE 'Ok'
END""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_staging LIMIT 20",conn))



print("--------------------------------------------------------------------")
print("---------------------CREATE FINAL CLEAN TABLE ----------------------")
conn.execute("DROP TABLE IF EXISTS orders_clean")
conn.execute("""CREATE TABLE orders_clean AS SELECT OrderID,
CustomerName,
Email,
Email_Valid,
Phone,
Phone_Valid,
City,
OrderDate,
Amount_Clean AS Amount,
Amount_Flag,
Status
FROM orders_staging""")
conn.commit()
print(pd.read_sql_query("SELECT * FROM orders_clean LIMIT 20",conn))


print("-------------------------------------------------------------------")
print("--------------------- Average order amount by city -----------------")
print(pd.read_sql_query("SELECT City,ROUND(AVG(Amount),2)AS Avg_Amount FROM orders_clean WHERE Amount_Flag='OK' GROUP BY City ORDER BY Amount DESC",conn))
conn.commit()


print("--------------------------------------------------------------------")
print("--------------------Total sales by city ----------------------------")
print(pd.read_sql_query("SELECT City,ROUND(SUM(Amount),2)AS Tot_Sales FROM orders_clean WHERE Amount_Flag='OK' GROUP BY City ORDER BY Tot_Sales DESC",conn))
conn.commit()


print("----------------------------------------------------------------------")
print("--------------------Number of orders by city -----------------------")
print(pd.read_sql_query("SELECT City,COUNT(*)AS Total_Orders FROM orders_clean GROUP BY City ORDER BY Total_Orders DESC",conn))
conn.commit()


print("--------------------------------------------------------------------")
print("------------------ Completed orders---------------------------------")
print(pd.read_sql_query("SELECT * FROM orders_clean WHERE Status='Completed'",conn))
print(pd.read_sql_query("SELECT COUNT(*)AS Cancelled_Orders FROM orders_clean WHERE Status='Cancelled'",conn))
print(pd.read_sql_query("SELECT City,COUNT(*)AS Cancelled_Orders FROM orders_clean WHERE Status='Cancelled'GROUP BY City ORDER BY Cancelled_Orders DESC",conn))
conn.commit()

print("---------------------------------------------------------------------")
print("--------------------Negative amount orders----------------------------")
print(pd.read_sql_query("SELECT OrderId,CustomerName,Amount,Status FROM orders_clean where Amount_Flag='Negative' ORDER BY Amount",conn))
conn.commit()

print("----------------------------------------------------------------------")
print("---------------- Monthly sales analysis-------------------------------")
print(pd.read_sql_query("""SELECT SUBSTR(OrderDate, 1, 7) AS Year_Month, COUNT(*) AS Total_Orders, ROUND(SUM(Amount), 2) AS Total_Sales FROM orders_clean WHERE
Amount_Flag = 'OK'
AND OrderDate IS NOT NULL
GROUP BY SUBSTR(OrderDate, 1, 7)
ORDER BY Year_Month""",conn))


print(pd.read_sql_query("SELECT * FROM orders_clean LIMIT 20",conn))


df.to_csv("orders_clean.csv", index=False)

conn.commit()
conn.close()
from google.colab import files

files.download("orders_clean.csv")
files.download("order_raw.db")

     
   tot_no_of_rows
0           10000
   OrderID        CustomerName                            Email  \
0     3105          Jaya Verma           jaya.verma801@mail.com   
1     6354        Farhan Yadav     farhan.yadav65@company.co.in   
2     8690       Chitra Reddy      chitra.reddy695company.co.in   
3     5858          Sneha Iyer      sneha.iyer945@company.co.in   
4     6012       kabir kumar         kabir.kumar305@outlook.com   
5     2653        Kabir Joshi            kabir.joshi17@mail.com   
6     5130  Deepa Chatterjee    deepa.chatterjee695@outlook.com   
7      416           Esha Iyer           esha.iyer391@gmail.com   
8     3539         Arun Yadav            arun.yadav764@mail.com   
9     1704       qadir singh            qadir.singh378@mail.com   

            Phone           City   OrderDate    Amount     Status  
0    733-307-7375          delhi  2024-01-05   3091.09    Pending  
1     23114 80745      Bengaluru  04-14-2023   1662.89   Refunded  
2    379-490-1038           pune  2024-08-03   1610.32   Refunded  
3  +91 8472931687   Coimbatore    11-08-2023   2265.32    Pending  
4      0400977061    Coimbatore   11-19-2024   2718.23  Cancelled  
5     75656 19112          Delhi  02-05-2023   ₹434.17    Pending  
6  +91 1238943682       CHENNAI   12-02-2024  2333.70    Refunded  
7     76927 43811     HYDERABAD   02-28-2025   3818.14    Pending  
8  +91 4798431264       Mumbai    02-05-2025   1297.01  Completed  
9      4311621065         Delhi   07-24-2025   2274.49    Pending  
-------------------------------------------------------------------
---------------------- Check duplicate OrderID values-------------- 
     OrderID  Duplicate_Count
0       9685                2
1       9606                2
2       9592                2
3       9569                2
4       9508                2
..       ...              ...
295      176                2
296      174                2
297      164                2
298       57                2
299       12                2

[300 rows x 2 columns]
-------------------------------------------------------------------
---------------------Check missing values -------------------------
   Missing_CustomerName  Missing_Email  Missing_Phone  Missing_City  \
0                   199            637            513          1011   

   Missing_OrderDate  Missing_Amount  Missing_Status  
0                329             255             839  
------------------------------------------------------------------
---------------------- Create a staging table----------------------
   Staging_Rows
0         10000
------------------------------------------------------------------
----------------------Remove exact duplicate rows ---------------
   Rows_After_Duplicate_Removal
0                          9700
----------------------------------------------------------------
----------------------Remove unwanted spaces -------------------
   OrderID      CustomerName                            Email           Phone  \
0     3105        Jaya Verma           jaya.verma801@mail.com    733-307-7375   
1     6354      Farhan Yadav     farhan.yadav65@company.co.in     23114 80745   
2     8690      Chitra Reddy     chitra.reddy695company.co.in    379-490-1038   
3     5858        Sneha Iyer      sneha.iyer945@company.co.in  +91 8472931687   
4     6012       kabir kumar       kabir.kumar305@outlook.com      0400977061   
5     2653       Kabir Joshi           kabir.joshi17@mail.com     75656 19112   
6     5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  +91 1238943682   
7      416         Esha Iyer           esha.iyer391@gmail.com     76927 43811   
8     3539        Arun Yadav           arun.yadav764@mail.com  +91 4798431264   
9     1704       qadir singh          qadir.singh378@mail.com      4311621065   

         City   OrderDate   Amount     Status  
0       delhi  2024-01-05  3091.09    Pending  
1   Bengaluru  04-14-2023  1662.89   Refunded  
2        pune  2024-08-03  1610.32   Refunded  
3  Coimbatore  11-08-2023  2265.32    Pending  
4  Coimbatore  11-19-2024  2718.23  Cancelled  
5       Delhi  02-05-2023  ₹434.17    Pending  
6     CHENNAI  12-02-2024  2333.70   Refunded  
7   HYDERABAD  02-28-2025  3818.14    Pending  
8      Mumbai  02-05-2025  1297.01  Completed  
9       Delhi  07-24-2025  2274.49    Pending  
----------------------------------------------------------------
-----------------Standardize Customer Names ----------------
   OrderID      CustomerName                            Email           Phone  \
0     3105        Jaya Verma           jaya.verma801@mail.com    733-307-7375   
1     6354      Farhan Yadav     farhan.yadav65@company.co.in     23114 80745   
2     8690      Chitra Reddy     chitra.reddy695company.co.in    379-490-1038   
3     5858        Sneha Iyer      sneha.iyer945@company.co.in  +91 8472931687   
4     6012       Kabir Kumar       kabir.kumar305@outlook.com      0400977061   
5     2653       Kabir Joshi           kabir.joshi17@mail.com     75656 19112   
6     5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  +91 1238943682   
7      416         Esha Iyer           esha.iyer391@gmail.com     76927 43811   
8     3539        Arun Yadav           arun.yadav764@mail.com  +91 4798431264   
9     1704       Qadir Singh          qadir.singh378@mail.com      4311621065   

         City   OrderDate   Amount     Status  
0       delhi  2024-01-05  3091.09    Pending  
1   Bengaluru  04-14-2023  1662.89   Refunded  
2        pune  2024-08-03  1610.32   Refunded  
3  Coimbatore  11-08-2023  2265.32    Pending  
4  Coimbatore  11-19-2024  2718.23  Cancelled  
5       Delhi  02-05-2023  ₹434.17    Pending  
6     CHENNAI  12-02-2024  2333.70   Refunded  
7   HYDERABAD  02-28-2025  3818.14    Pending  
8      Mumbai  02-05-2025  1297.01  Completed  
9       Delhi  07-24-2025  2274.49    Pending  
-----------------------------------------------------------------
----------------Standardize City names --------------------------
   OrderID      CustomerName                            Email           Phone  \
0     3105        Jaya Verma           jaya.verma801@mail.com    733-307-7375   
1     6354      Farhan Yadav     farhan.yadav65@company.co.in     23114 80745   
2     8690      Chitra Reddy     chitra.reddy695company.co.in    379-490-1038   
3     5858        Sneha Iyer      sneha.iyer945@company.co.in  +91 8472931687   
4     6012       Kabir Kumar       kabir.kumar305@outlook.com      0400977061   
5     2653       Kabir Joshi           kabir.joshi17@mail.com     75656 19112   
6     5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  +91 1238943682   
7      416         Esha Iyer           esha.iyer391@gmail.com     76927 43811   
8     3539        Arun Yadav           arun.yadav764@mail.com  +91 4798431264   
9     1704       Qadir Singh          qadir.singh378@mail.com      4311621065   

         City   OrderDate   Amount     Status  
0       Delhi  2024-01-05  3091.09    Pending  
1   Bengaluru  04-14-2023  1662.89   Refunded  
2        Pune  2024-08-03  1610.32   Refunded  
3  Coimbatore  11-08-2023  2265.32    Pending  
4  Coimbatore  11-19-2024  2718.23  Cancelled  
5       Delhi  02-05-2023  ₹434.17    Pending  
6     Chennai  12-02-2024  2333.70   Refunded  
7   Hyderabad  02-28-2025  3818.14    Pending  
8      Mumbai  02-05-2025  1297.01  Completed  
9       Delhi  07-24-2025  2274.49    Pending  
-----------------------------------------------------------------
------------------Convert Email to lowercase --------------------
   OrderID      CustomerName                            Email           Phone  \
0     3105        Jaya Verma           jaya.verma801@mail.com    733-307-7375   
1     6354      Farhan Yadav     farhan.yadav65@company.co.in     23114 80745   
2     8690      Chitra Reddy     chitra.reddy695company.co.in    379-490-1038   
3     5858        Sneha Iyer      sneha.iyer945@company.co.in  +91 8472931687   
4     6012       Kabir Kumar       kabir.kumar305@outlook.com      0400977061   
5     2653       Kabir Joshi           kabir.joshi17@mail.com     75656 19112   
6     5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  +91 1238943682   
7      416         Esha Iyer           esha.iyer391@gmail.com     76927 43811   
8     3539        Arun Yadav           arun.yadav764@mail.com  +91 4798431264   
9     1704       Qadir Singh          qadir.singh378@mail.com      4311621065   

         City   OrderDate   Amount     Status  
0       Delhi  2024-01-05  3091.09    Pending  
1   Bengaluru  04-14-2023  1662.89   Refunded  
2        Pune  2024-08-03  1610.32   Refunded  
3  Coimbatore  11-08-2023  2265.32    Pending  
4  Coimbatore  11-19-2024  2718.23  Cancelled  
5       Delhi  02-05-2023  ₹434.17    Pending  
6     Chennai  12-02-2024  2333.70   Refunded  
7   Hyderabad  02-28-2025  3818.14    Pending  
8      Mumbai  02-05-2025  1297.01  Completed  
9       Delhi  07-24-2025  2274.49    Pending  
-----------------------------------------------------------------
---------------------Standardize Status -------------------------
   OrderID      CustomerName                            Email           Phone  \
0     3105        Jaya Verma           jaya.verma801@mail.com    733-307-7375   
1     6354      Farhan Yadav     farhan.yadav65@company.co.in     23114 80745   
2     8690      Chitra Reddy     chitra.reddy695company.co.in    379-490-1038   
3     5858        Sneha Iyer      sneha.iyer945@company.co.in  +91 8472931687   
4     6012       Kabir Kumar       kabir.kumar305@outlook.com      0400977061   
5     2653       Kabir Joshi           kabir.joshi17@mail.com     75656 19112   
6     5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  +91 1238943682   
7      416         Esha Iyer           esha.iyer391@gmail.com     76927 43811   
8     3539        Arun Yadav           arun.yadav764@mail.com  +91 4798431264   
9     1704       Qadir Singh          qadir.singh378@mail.com      4311621065   

         City   OrderDate   Amount     Status  
0       Delhi  2024-01-05  3091.09    Pending  
1   Bengaluru  04-14-2023  1662.89   Refunded  
2        Pune  2024-08-03  1610.32   Refunded  
3  Coimbatore  11-08-2023  2265.32    Pending  
4  Coimbatore  11-19-2024  2718.23  Cancelled  
5       Delhi  02-05-2023  ₹434.17    Pending  
6     Chennai  12-02-2024  2333.70   Refunded  
7   Hyderabad  02-28-2025  3818.14    Pending  
8      Mumbai  02-05-2025  1297.01  Completed  
9       Delhi  07-24-2025  2274.49    Pending  
------------------------------------------------------------------
      Status   tot
0   Refunded  2311
1    Pending  2182
2  Completed  2180
3  Cancelled  2164
4       None   804
5   Canceled    59
-------------------------------------------------------------------
--------------------Handle missing values -------------------------
    OrderID      CustomerName                            Email  \
0      3105        Jaya Verma           jaya.verma801@mail.com   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in   
2      8690      Chitra Reddy     chitra.reddy695company.co.in   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com   
5      2653       Kabir Joshi           kabir.joshi17@mail.com   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com   
7       416         Esha Iyer           esha.iyer391@gmail.com   
8      3539        Arun Yadav           arun.yadav764@mail.com   
9      1704       Qadir Singh          qadir.singh378@mail.com   
10     3528       Wasim Gupta          wasim.gupta712@mail.com   
11     7900        Esha Joshi          esha.joshi714@yahoo.com   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com   
13      747         Hari Khan           hari.khan234@yahoo.com   
14     2818       Manoj Verma         manoj.verma369@gmail.com   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com   
17     3328          Zoya Das                 zoya.das723@mail   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com   

             Phone        City   OrderDate     Amount     Status  
0     733-307-7375       Delhi  2024-01-05    3091.09    Pending  
1      23114 80745   Bengaluru  04-14-2023    1662.89   Refunded  
2     379-490-1038        Pune  2024-08-03    1610.32   Refunded  
3   +91 8472931687  Coimbatore  11-08-2023    2265.32    Pending  
4       0400977061  Coimbatore  11-19-2024    2718.23  Cancelled  
5      75656 19112       Delhi  02-05-2023    ₹434.17    Pending  
6   +91 1238943682     Chennai  12-02-2024    2333.70   Refunded  
7      76927 43811   Hyderabad  02-28-2025    3818.14    Pending  
8   +91 4798431264      Mumbai  02-05-2025    1297.01  Completed  
9       4311621065       Delhi  07-24-2025    2274.49    Pending  
10  (162) 598-3913     Chennai  08/03/2023    1928.67  Completed  
11  (520) 522-7568     Chennai  2025-06-20    2584.67   Refunded  
12  (626) 959-5197      Mumbai  14/08/2024    3557.42  Completed  
13    442-347-1998     Chennai  06-27-2023    3198.56  Cancelled  
14      7293060898     Unknown  03-19-2025    3910.42    Pending  
15      6269877351   Hyderabad  07-17-2024    4958.93    Unknown  
16  (295) 414-5386   Bengaluru  09-24-2025  ₹1,405.41  Completed  
17    848-550-5635     Chennai  2025-05-03       None   Refunded  
18    570-009-0038       Delhi  2025-02-01    3580.75  Completed  
19     67380 07506  Coimbatore  10-08-2023     2729.8    Pending  
-------------------------------------------------------------------
----------------------Standardize OrderDate------------------------
    OrderID      CustomerName                            Email  \
0      3105        Jaya Verma           jaya.verma801@mail.com   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in   
2      8690      Chitra Reddy     chitra.reddy695company.co.in   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com   
5      2653       Kabir Joshi           kabir.joshi17@mail.com   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com   
7       416         Esha Iyer           esha.iyer391@gmail.com   
8      3539        Arun Yadav           arun.yadav764@mail.com   
9      1704       Qadir Singh          qadir.singh378@mail.com   
10     3528       Wasim Gupta          wasim.gupta712@mail.com   
11     7900        Esha Joshi          esha.joshi714@yahoo.com   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com   
13      747         Hari Khan           hari.khan234@yahoo.com   
14     2818       Manoj Verma         manoj.verma369@gmail.com   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com   
17     3328          Zoya Das                 zoya.das723@mail   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com   

             Phone        City   OrderDate     Amount     Status  
0     733-307-7375       Delhi  2024-01-05    3091.09    Pending  
1      23114 80745   Bengaluru  2023-04-14    1662.89   Refunded  
2     379-490-1038        Pune  2024-08-03    1610.32   Refunded  
3   +91 8472931687  Coimbatore  2023-11-08    2265.32    Pending  
4       0400977061  Coimbatore  2024-11-19    2718.23  Cancelled  
5      75656 19112       Delhi  2023-02-05    ₹434.17    Pending  
6   +91 1238943682     Chennai  2024-12-02    2333.70   Refunded  
7      76927 43811   Hyderabad  2025-02-28    3818.14    Pending  
8   +91 4798431264      Mumbai  2025-02-05    1297.01  Completed  
9       4311621065       Delhi  2025-07-24    2274.49    Pending  
10  (162) 598-3913     Chennai        None    1928.67  Completed  
11  (520) 522-7568     Chennai  2025-06-20    2584.67   Refunded  
12  (626) 959-5197      Mumbai        None    3557.42  Completed  
13    442-347-1998     Chennai  2023-06-27    3198.56  Cancelled  
14      7293060898     Unknown  2025-03-19    3910.42    Pending  
15      6269877351   Hyderabad  2024-07-17    4958.93    Unknown  
16  (295) 414-5386   Bengaluru  2025-09-24  ₹1,405.41  Completed  
17    848-550-5635     Chennai  2025-05-03       None   Refunded  
18    570-009-0038       Delhi  2025-02-01    3580.75  Completed  
19     67380 07506  Coimbatore  2023-10-08     2729.8    Pending  
-------------------------------------------------------------------
------------Validate Email ----------------------------------------
    OrderID      CustomerName                            Email  \
0      3105        Jaya Verma           jaya.verma801@mail.com   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in   
2      8690      Chitra Reddy     chitra.reddy695company.co.in   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com   
5      2653       Kabir Joshi           kabir.joshi17@mail.com   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com   
7       416         Esha Iyer           esha.iyer391@gmail.com   
8      3539        Arun Yadav           arun.yadav764@mail.com   
9      1704       Qadir Singh          qadir.singh378@mail.com   
10     3528       Wasim Gupta          wasim.gupta712@mail.com   
11     7900        Esha Joshi          esha.joshi714@yahoo.com   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com   
13      747         Hari Khan           hari.khan234@yahoo.com   
14     2818       Manoj Verma         manoj.verma369@gmail.com   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com   
17     3328          Zoya Das                 zoya.das723@mail   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com   

             Phone        City   OrderDate     Amount     Status  Email_Valid  
0     733-307-7375       Delhi  2024-01-05    3091.09    Pending            1  
1      23114 80745   Bengaluru  2023-04-14    1662.89   Refunded            1  
2     379-490-1038        Pune  2024-08-03    1610.32   Refunded            0  
3   +91 8472931687  Coimbatore  2023-11-08    2265.32    Pending            1  
4       0400977061  Coimbatore  2024-11-19    2718.23  Cancelled            1  
5      75656 19112       Delhi  2023-02-05    ₹434.17    Pending            1  
6   +91 1238943682     Chennai  2024-12-02    2333.70   Refunded            1  
7      76927 43811   Hyderabad  2025-02-28    3818.14    Pending            1  
8   +91 4798431264      Mumbai  2025-02-05    1297.01  Completed            1  
9       4311621065       Delhi  2025-07-24    2274.49    Pending            1  
10  (162) 598-3913     Chennai        None    1928.67  Completed            1  
11  (520) 522-7568     Chennai  2025-06-20    2584.67   Refunded            1  
12  (626) 959-5197      Mumbai        None    3557.42  Completed            1  
13    442-347-1998     Chennai  2023-06-27    3198.56  Cancelled            1  
14      7293060898     Unknown  2025-03-19    3910.42    Pending            1  
15      6269877351   Hyderabad  2024-07-17    4958.93    Unknown            1  
16  (295) 414-5386   Bengaluru  2025-09-24  ₹1,405.41  Completed            1  
17    848-550-5635     Chennai  2025-05-03       None   Refunded            0  
18    570-009-0038       Delhi  2025-02-01    3580.75  Completed            1  
19     67380 07506  Coimbatore  2023-10-08     2729.8    Pending            1  
--------------------------------------------------------------------
-----------------Clean Phone numbers -------------------------------
    OrderID      CustomerName                            Email         Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com    7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in    2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in    3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  918472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com    0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com    7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  911238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com    7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  914798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com    4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com    1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com    5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com    6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com    4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com    7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in    6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com    2954145386   
17     3328          Zoya Das                 zoya.das723@mail    8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com    5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com    6738007506   

          City   OrderDate     Amount     Status  Email_Valid  
0        Delhi  2024-01-05    3091.09    Pending            1  
1    Bengaluru  2023-04-14    1662.89   Refunded            1  
2         Pune  2024-08-03    1610.32   Refunded            0  
3   Coimbatore  2023-11-08    2265.32    Pending            1  
4   Coimbatore  2024-11-19    2718.23  Cancelled            1  
5        Delhi  2023-02-05    ₹434.17    Pending            1  
6      Chennai  2024-12-02    2333.70   Refunded            1  
7    Hyderabad  2025-02-28    3818.14    Pending            1  
8       Mumbai  2025-02-05    1297.01  Completed            1  
9        Delhi  2025-07-24    2274.49    Pending            1  
10     Chennai        None    1928.67  Completed            1  
11     Chennai  2025-06-20    2584.67   Refunded            1  
12      Mumbai        None    3557.42  Completed            1  
13     Chennai  2023-06-27    3198.56  Cancelled            1  
14     Unknown  2025-03-19    3910.42    Pending            1  
15   Hyderabad  2024-07-17    4958.93    Unknown            1  
16   Bengaluru  2025-09-24  ₹1,405.41  Completed            1  
17     Chennai  2025-05-03       None   Refunded            0  
18       Delhi  2025-02-01    3580.75  Completed            1  
19  Coimbatore  2023-10-08     2729.8    Pending            1  
-------------------------------------------------------------------
-----------------Remove +91 country code --------------------------
    OrderID      CustomerName                            Email       Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com  7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in  2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in  3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  8472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com  0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com  7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  1238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com  7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  4798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com  4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com  1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com  5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com  6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com  4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com  7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in  6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com  2954145386   
17     3328          Zoya Das                 zoya.das723@mail  8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com  5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com  6738007506   

          City   OrderDate     Amount     Status  Email_Valid  
0        Delhi  2024-01-05    3091.09    Pending            1  
1    Bengaluru  2023-04-14    1662.89   Refunded            1  
2         Pune  2024-08-03    1610.32   Refunded            0  
3   Coimbatore  2023-11-08    2265.32    Pending            1  
4   Coimbatore  2024-11-19    2718.23  Cancelled            1  
5        Delhi  2023-02-05    ₹434.17    Pending            1  
6      Chennai  2024-12-02    2333.70   Refunded            1  
7    Hyderabad  2025-02-28    3818.14    Pending            1  
8       Mumbai  2025-02-05    1297.01  Completed            1  
9        Delhi  2025-07-24    2274.49    Pending            1  
10     Chennai        None    1928.67  Completed            1  
11     Chennai  2025-06-20    2584.67   Refunded            1  
12      Mumbai        None    3557.42  Completed            1  
13     Chennai  2023-06-27    3198.56  Cancelled            1  
14     Unknown  2025-03-19    3910.42    Pending            1  
15   Hyderabad  2024-07-17    4958.93    Unknown            1  
16   Bengaluru  2025-09-24  ₹1,405.41  Completed            1  
17     Chennai  2025-05-03       None   Refunded            0  
18       Delhi  2025-02-01    3580.75  Completed            1  
19  Coimbatore  2023-10-08     2729.8    Pending            1  
-------------------------------------------------------------------
------------------Validate Phone numbers --------------------------
    OrderID      CustomerName                            Email       Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com  7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in  2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in  3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  8472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com  0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com  7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  1238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com  7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  4798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com  4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com  1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com  5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com  6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com  4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com  7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in  6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com  2954145386   
17     3328          Zoya Das                 zoya.das723@mail  8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com  5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com  6738007506   

          City   OrderDate     Amount     Status  Email_Valid  Phone_Valid  
0        Delhi  2024-01-05    3091.09    Pending            1            1  
1    Bengaluru  2023-04-14    1662.89   Refunded            1            1  
2         Pune  2024-08-03    1610.32   Refunded            0            1  
3   Coimbatore  2023-11-08    2265.32    Pending            1            1  
4   Coimbatore  2024-11-19    2718.23  Cancelled            1            1  
5        Delhi  2023-02-05    ₹434.17    Pending            1            1  
6      Chennai  2024-12-02    2333.70   Refunded            1            1  
7    Hyderabad  2025-02-28    3818.14    Pending            1            1  
8       Mumbai  2025-02-05    1297.01  Completed            1            1  
9        Delhi  2025-07-24    2274.49    Pending            1            1  
10     Chennai        None    1928.67  Completed            1            1  
11     Chennai  2025-06-20    2584.67   Refunded            1            1  
12      Mumbai        None    3557.42  Completed            1            1  
13     Chennai  2023-06-27    3198.56  Cancelled            1            1  
14     Unknown  2025-03-19    3910.42    Pending            1            1  
15   Hyderabad  2024-07-17    4958.93    Unknown            1            1  
16   Bengaluru  2025-09-24  ₹1,405.41  Completed            1            1  
17     Chennai  2025-05-03       None   Refunded            0            1  
18       Delhi  2025-02-01    3580.75  Completed            1            1  
19  Coimbatore  2023-10-08     2729.8    Pending            1            1  
-------------------------------------------------------------------
--------------------Clean Amount ----------------------------------
    OrderID      CustomerName                            Email       Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com  7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in  2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in  3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  8472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com  0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com  7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  1238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com  7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  4798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com  4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com  1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com  5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com  6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com  4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com  7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in  6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com  2954145386   
17     3328          Zoya Das                 zoya.das723@mail  8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com  5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com  6738007506   

          City   OrderDate   Amount     Status  Email_Valid  Phone_Valid  
0        Delhi  2024-01-05  3091.09    Pending            1            1  
1    Bengaluru  2023-04-14  1662.89   Refunded            1            1  
2         Pune  2024-08-03  1610.32   Refunded            0            1  
3   Coimbatore  2023-11-08  2265.32    Pending            1            1  
4   Coimbatore  2024-11-19  2718.23  Cancelled            1            1  
5        Delhi  2023-02-05   434.17    Pending            1            1  
6      Chennai  2024-12-02  2333.70   Refunded            1            1  
7    Hyderabad  2025-02-28  3818.14    Pending            1            1  
8       Mumbai  2025-02-05  1297.01  Completed            1            1  
9        Delhi  2025-07-24  2274.49    Pending            1            1  
10     Chennai        None  1928.67  Completed            1            1  
11     Chennai  2025-06-20  2584.67   Refunded            1            1  
12      Mumbai        None  3557.42  Completed            1            1  
13     Chennai  2023-06-27  3198.56  Cancelled            1            1  
14     Unknown  2025-03-19  3910.42    Pending            1            1  
15   Hyderabad  2024-07-17  4958.93    Unknown            1            1  
16   Bengaluru  2025-09-24  1405.41  Completed            1            1  
17     Chennai  2025-05-03     None   Refunded            0            1  
18       Delhi  2025-02-01  3580.75  Completed            1            1  
19  Coimbatore  2023-10-08   2729.8    Pending            1            1  
-------------------------------------------------------------------
-------------------Create numeric Amount column -------------------
    OrderID      CustomerName                            Email       Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com  7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in  2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in  3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  8472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com  0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com  7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  1238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com  7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  4798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com  4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com  1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com  5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com  6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com  4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com  7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in  6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com  2954145386   
17     3328          Zoya Das                 zoya.das723@mail  8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com  5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com  6738007506   

          City   OrderDate   Amount     Status  Email_Valid  Phone_Valid  \
0        Delhi  2024-01-05  3091.09    Pending            1            1   
1    Bengaluru  2023-04-14  1662.89   Refunded            1            1   
2         Pune  2024-08-03  1610.32   Refunded            0            1   
3   Coimbatore  2023-11-08  2265.32    Pending            1            1   
4   Coimbatore  2024-11-19  2718.23  Cancelled            1            1   
5        Delhi  2023-02-05   434.17    Pending            1            1   
6      Chennai  2024-12-02  2333.70   Refunded            1            1   
7    Hyderabad  2025-02-28  3818.14    Pending            1            1   
8       Mumbai  2025-02-05  1297.01  Completed            1            1   
9        Delhi  2025-07-24  2274.49    Pending            1            1   
10     Chennai        None  1928.67  Completed            1            1   
11     Chennai  2025-06-20  2584.67   Refunded            1            1   
12      Mumbai        None  3557.42  Completed            1            1   
13     Chennai  2023-06-27  3198.56  Cancelled            1            1   
14     Unknown  2025-03-19  3910.42    Pending            1            1   
15   Hyderabad  2024-07-17  4958.93    Unknown            1            1   
16   Bengaluru  2025-09-24  1405.41  Completed            1            1   
17     Chennai  2025-05-03     None   Refunded            0            1   
18       Delhi  2025-02-01  3580.75  Completed            1            1   
19  Coimbatore  2023-10-08   2729.8    Pending            1            1   

    Amount_Clean  
0        3091.09  
1        1662.89  
2        1610.32  
3        2265.32  
4        2718.23  
5         434.17  
6        2333.70  
7        3818.14  
8        1297.01  
9        2274.49  
10       1928.67  
11       2584.67  
12       3557.42  
13       3198.56  
14       3910.42  
15       4958.93  
16       1405.41  
17           NaN  
18       3580.75  
19       2729.80  
-------------------------------------------------------------------
--------------------- Identify Amount problems --------------------
    OrderID      CustomerName                            Email       Phone  \
0      3105        Jaya Verma           jaya.verma801@mail.com  7333077375   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in  2311480745   
2      8690      Chitra Reddy     chitra.reddy695company.co.in  3794901038   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in  8472931687   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com  0400977061   
5      2653       Kabir Joshi           kabir.joshi17@mail.com  7565619112   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com  1238943682   
7       416         Esha Iyer           esha.iyer391@gmail.com  7692743811   
8      3539        Arun Yadav           arun.yadav764@mail.com  4798431264   
9      1704       Qadir Singh          qadir.singh378@mail.com  4311621065   
10     3528       Wasim Gupta          wasim.gupta712@mail.com  1625983913   
11     7900        Esha Joshi          esha.joshi714@yahoo.com  5205227568   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com  6269595197   
13      747         Hari Khan           hari.khan234@yahoo.com  4423471998   
14     2818       Manoj Verma         manoj.verma369@gmail.com  7293060898   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in  6269877351   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com  2954145386   
17     3328          Zoya Das                 zoya.das723@mail  8485505635   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com  5700090038   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com  6738007506   

          City   OrderDate   Amount     Status  Email_Valid  Phone_Valid  \
0        Delhi  2024-01-05  3091.09    Pending            1            1   
1    Bengaluru  2023-04-14  1662.89   Refunded            1            1   
2         Pune  2024-08-03  1610.32   Refunded            0            1   
3   Coimbatore  2023-11-08  2265.32    Pending            1            1   
4   Coimbatore  2024-11-19  2718.23  Cancelled            1            1   
5        Delhi  2023-02-05   434.17    Pending            1            1   
6      Chennai  2024-12-02  2333.70   Refunded            1            1   
7    Hyderabad  2025-02-28  3818.14    Pending            1            1   
8       Mumbai  2025-02-05  1297.01  Completed            1            1   
9        Delhi  2025-07-24  2274.49    Pending            1            1   
10     Chennai        None  1928.67  Completed            1            1   
11     Chennai  2025-06-20  2584.67   Refunded            1            1   
12      Mumbai        None  3557.42  Completed            1            1   
13     Chennai  2023-06-27  3198.56  Cancelled            1            1   
14     Unknown  2025-03-19  3910.42    Pending            1            1   
15   Hyderabad  2024-07-17  4958.93    Unknown            1            1   
16   Bengaluru  2025-09-24  1405.41  Completed            1            1   
17     Chennai  2025-05-03     None   Refunded            0            1   
18       Delhi  2025-02-01  3580.75  Completed            1            1   
19  Coimbatore  2023-10-08   2729.8    Pending            1            1   

    Amount_Clean Amount_Flag  
0        3091.09          Ok  
1        1662.89          Ok  
2        1610.32          Ok  
3        2265.32          Ok  
4        2718.23          Ok  
5         434.17          Ok  
6        2333.70          Ok  
7        3818.14          Ok  
8        1297.01          Ok  
9        2274.49          Ok  
10       1928.67          Ok  
11       2584.67          Ok  
12       3557.42          Ok  
13       3198.56          Ok  
14       3910.42          Ok  
15       4958.93          Ok  
16       1405.41          Ok  
17           NaN     Missing  
18       3580.75          Ok  
19       2729.80          Ok  
--------------------------------------------------------------------
---------------------CREATE FINAL CLEAN TABLE ----------------------
    OrderID      CustomerName                            Email  Email_Valid  \
0      3105        Jaya Verma           jaya.verma801@mail.com            1   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in            1   
2      8690      Chitra Reddy     chitra.reddy695company.co.in            0   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in            1   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com            1   
5      2653       Kabir Joshi           kabir.joshi17@mail.com            1   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com            1   
7       416         Esha Iyer           esha.iyer391@gmail.com            1   
8      3539        Arun Yadav           arun.yadav764@mail.com            1   
9      1704       Qadir Singh          qadir.singh378@mail.com            1   
10     3528       Wasim Gupta          wasim.gupta712@mail.com            1   
11     7900        Esha Joshi          esha.joshi714@yahoo.com            1   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com            1   
13      747         Hari Khan           hari.khan234@yahoo.com            1   
14     2818       Manoj Verma         manoj.verma369@gmail.com            1   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in            1   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com            1   
17     3328          Zoya Das                 zoya.das723@mail            0   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com            1   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com            1   

         Phone  Phone_Valid        City   OrderDate   Amount Amount_Flag  \
0   7333077375            1       Delhi  2024-01-05  3091.09          Ok   
1   2311480745            1   Bengaluru  2023-04-14  1662.89          Ok   
2   3794901038            1        Pune  2024-08-03  1610.32          Ok   
3   8472931687            1  Coimbatore  2023-11-08  2265.32          Ok   
4   0400977061            1  Coimbatore  2024-11-19  2718.23          Ok   
5   7565619112            1       Delhi  2023-02-05   434.17          Ok   
6   1238943682            1     Chennai  2024-12-02  2333.70          Ok   
7   7692743811            1   Hyderabad  2025-02-28  3818.14          Ok   
8   4798431264            1      Mumbai  2025-02-05  1297.01          Ok   
9   4311621065            1       Delhi  2025-07-24  2274.49          Ok   
10  1625983913            1     Chennai        None  1928.67          Ok   
11  5205227568            1     Chennai  2025-06-20  2584.67          Ok   
12  6269595197            1      Mumbai        None  3557.42          Ok   
13  4423471998            1     Chennai  2023-06-27  3198.56          Ok   
14  7293060898            1     Unknown  2025-03-19  3910.42          Ok   
15  6269877351            1   Hyderabad  2024-07-17  4958.93          Ok   
16  2954145386            1   Bengaluru  2025-09-24  1405.41          Ok   
17  8485505635            1     Chennai  2025-05-03      NaN     Missing   
18  5700090038            1       Delhi  2025-02-01  3580.75          Ok   
19  6738007506            1  Coimbatore  2023-10-08  2729.80          Ok   

       Status  
0     Pending  
1    Refunded  
2    Refunded  
3     Pending  
4   Cancelled  
5     Pending  
6    Refunded  
7     Pending  
8   Completed  
9     Pending  
10  Completed  
11   Refunded  
12  Completed  
13  Cancelled  
14    Pending  
15    Unknown  
16  Completed  
17   Refunded  
18  Completed  
19    Pending  
-------------------------------------------------------------------
--------------------- Average order amount by city -----------------
Empty DataFrame
Columns: [City, Avg_Amount]
Index: []
--------------------------------------------------------------------
--------------------Total sales by city ----------------------------
Empty DataFrame
Columns: [City, Tot_Sales]
Index: []
----------------------------------------------------------------------
--------------------Number of orders by city -----------------------
         City  Total_Orders
0  Coimbatore          1149
1     Chennai          1109
2       Delhi          1105
3      Mumbai          1085
4        Pune          1081
5     Kolkata          1080
6   Bengaluru          1060
7   Hyderabad          1049
8     Unknown           982
--------------------------------------------------------------------
------------------ Completed orders---------------------------------
      OrderID   CustomerName                         Email  Email_Valid  \
0        3539     Arun Yadav        arun.yadav764@mail.com            1   
1        3528    Wasim Gupta       wasim.gupta712@mail.com            1   
2        3027      Ravi Iyer      ravi.iyer314@outlook.com            1   
3        6985    Kabir Mehta      kabir.mehta681@yahoo.com            1   
4        7954  Bhavna Pillai    bhavna.pillai560@gmail.com            1   
...       ...            ...                           ...          ...   
2175     3886   Wasim Pillai  wasim.pillai48@company.co.in            1   
2176     7934    Hari Kapoor  hari.kapoor960@company.co.in            1   
2177     5725     Xena Kumar       xena.kumar267@yahoo.com            1   
2178     6766    Kabir Joshi    kabir.joshi718@outlook.com            1   
2179     7260    Omkar Verma                          None            0   

           Phone  Phone_Valid        City   OrderDate   Amount Amount_Flag  \
0     4798431264            1      Mumbai  2025-02-05  1297.01          Ok   
1     1625983913            1     Chennai        None  1928.67          Ok   
2     6269595197            1      Mumbai        None  3557.42          Ok   
3     2954145386            1   Bengaluru  2025-09-24  1405.41          Ok   
4     5700090038            1       Delhi  2025-02-01  3580.75          Ok   
...          ...          ...         ...         ...      ...         ...   
2175  4094403684            1  Coimbatore        None  3163.29          Ok   
2176  7217767107            1  Coimbatore  2025-01-16  2450.31          Ok   
2177  7065351239            1     Kolkata  2023-09-05  1011.95          Ok   
2178  4854857216            1     Chennai  2024-02-16  3312.43          Ok   
2179  4139331405            1      Mumbai  2025-06-05  1938.64          Ok   

         Status  
0     Completed  
1     Completed  
2     Completed  
3     Completed  
4     Completed  
...         ...  
2175  Completed  
2176  Completed  
2177  Completed  
2178  Completed  
2179  Completed  

[2180 rows x 11 columns]
   Cancelled_Orders
0              2164
         City  Cancelled_Orders
0     Chennai               268
1  Coimbatore               254
2      Mumbai               252
3        Pune               242
4     Kolkata               241
5     Unknown               239
6   Hyderabad               231
7       Delhi               226
8   Bengaluru               211
---------------------------------------------------------------------
--------------------Negative amount orders----------------------------
Empty DataFrame
Columns: [OrderID, CustomerName, Amount, Status]
Index: []
----------------------------------------------------------------------
---------------- Monthly sales analysis-------------------------------
Empty DataFrame
Columns: [Year_Month, Total_Orders, Total_Sales]
Index: []
    OrderID      CustomerName                            Email  Email_Valid  \
0      3105        Jaya Verma           jaya.verma801@mail.com            1   
1      6354      Farhan Yadav     farhan.yadav65@company.co.in            1   
2      8690      Chitra Reddy     chitra.reddy695company.co.in            0   
3      5858        Sneha Iyer      sneha.iyer945@company.co.in            1   
4      6012       Kabir Kumar       kabir.kumar305@outlook.com            1   
5      2653       Kabir Joshi           kabir.joshi17@mail.com            1   
6      5130  Deepa Chatterjee  deepa.chatterjee695@outlook.com            1   
7       416         Esha Iyer           esha.iyer391@gmail.com            1   
8      3539        Arun Yadav           arun.yadav764@mail.com            1   
9      1704       Qadir Singh          qadir.singh378@mail.com            1   
10     3528       Wasim Gupta          wasim.gupta712@mail.com            1   
11     7900        Esha Joshi          esha.joshi714@yahoo.com            1   
12     3027         Ravi Iyer         ravi.iyer314@outlook.com            1   
13      747         Hari Khan           hari.khan234@yahoo.com            1   
14     2818       Manoj Verma         manoj.verma369@gmail.com            1   
15     1916     Vikram Pillai   vikram.pillai568@company.co.in            1   
16     6985       Kabir Mehta         kabir.mehta681@yahoo.com            1   
17     3328          Zoya Das                 zoya.das723@mail            0   
18     7954     Bhavna Pillai       bhavna.pillai560@gmail.com            1   
19     1770       Omkar Gupta       omkar.gupta615@outlook.com            1   

         Phone  Phone_Valid        City   OrderDate   Amount Amount_Flag  \
0   7333077375            1       Delhi  2024-01-05  3091.09          Ok   
1   2311480745            1   Bengaluru  2023-04-14  1662.89          Ok   
2   3794901038            1        Pune  2024-08-03  1610.32          Ok   
3   8472931687            1  Coimbatore  2023-11-08  2265.32          Ok   
4   0400977061            1  Coimbatore  2024-11-19  2718.23          Ok   
5   7565619112            1       Delhi  2023-02-05   434.17          Ok   
6   1238943682            1     Chennai  2024-12-02  2333.70          Ok   
7   7692743811            1   Hyderabad  2025-02-28  3818.14          Ok   
8   4798431264            1      Mumbai  2025-02-05  1297.01          Ok   
9   4311621065            1       Delhi  2025-07-24  2274.49          Ok   
10  1625983913            1     Chennai        None  1928.67          Ok   
11  5205227568            1     Chennai  2025-06-20  2584.67          Ok   
12  6269595197            1      Mumbai        None  3557.42          Ok   
13  4423471998            1     Chennai  2023-06-27  3198.56          Ok   
14  7293060898            1     Unknown  2025-03-19  3910.42          Ok   
15  6269877351            1   Hyderabad  2024-07-17  4958.93          Ok   
16  2954145386            1   Bengaluru  2025-09-24  1405.41          Ok   
17  8485505635            1     Chennai  2025-05-03      NaN     Missing   
18  5700090038            1       Delhi  2025-02-01  3580.75          Ok   
19  6738007506            1  Coimbatore  2023-10-08  2729.80          Ok   

       Status  
0     Pending  
1    Refunded  
2    Refunded  
3     Pending  
4   Cancelled  
5     Pending  
6    Refunded  
7     Pending  
8   Completed  
9     Pending  
10  Completed  
11   Refunded  
12  Completed  
13  Cancelled  
14    Pending  
15    Unknown  
16  Completed  
17   Refunded  
18  Completed  
19    Pending  
