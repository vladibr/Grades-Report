USE Ruppin_Students
GO

--Veriables
DECLARE @avg float
DECLARE @number_of_students int
GO

--Students table
CREATE TABLE students
(
username_id char(20) not null,
CONSTRAINT PK_username_id PRIMARY KEY CLUSTERED (username_id),
password nvarchar(200) not null,
email nvarchar(50) 
)

--Temp table for our grades
CREATE TABLE temp_table
(
[code] nvarchar(100),
[subject] nvarchar(100),
[type] nvarchar(100),
[test] nvarchar(100),
[grade] nvarchar(100)
)

--All Courses Table
CREATE TABLE courses
(
[user_id] nvarchar(100) not null,
[code] nvarchar(100),
[subject] nvarchar(100),
[type] nvarchar(100),
[test] nvarchar(100),
[grade] nvarchar(100)
)
GO

--add user if not exist
CREATE PROC add_user @user_id nvarchar(20), @user_pass nvarchar(200), @email nvarchar(50)
AS
IF EXISTS (SELECT * FROM [dbo].[students] 
		   WHERE [username_id] = @user_id AND [password] = @user_pass AND [email] != @email)
BEGIN
UPDATE [students] SET [email] = @email WHERE [username_id] = @user_id
END

ELSE IF NOT EXISTS (SELECT * FROM [dbo].[students]
               WHERE [username_id] = @user_id)
BEGIN
INSERT [students] ([username_id], [password], [email]) VALUES (@user_id, @user_pass, @email)
END
GO

--Returns number of students
CREATE PROC number_of_students
AS
DECLARE @no_students int
SELECT @no_students = COUNT(DISTINCT [user_id])
FROM [courses]
RETURN @no_students
GO

--Return user grades average
CREATE PROC user_average @user_id nvarchar(30)
AS
DECLARE @avg float
SELECT @avg = avg(convert(float, [grade]))
FROM courses
WHERE isnumeric([grade])=1 AND [user_id] = @user_id
RETURN @avg
GO

--Proc that add all user grades to grades table
ALTER PROC insert_grades @file_name nvarchar(100)
AS
DELETE FROM [temp_table]

DECLARE @sql NVARCHAR(4000) = '
BULK INSERT temp_table
FROM ''D:\Salman_Api\Grades_Report_Query\Reports\' + @file_name + '.csv''
WITH 
    ( 
    CODEPAGE = ''ACP'',  
    FIRSTROW = 3 ,  
    MAXERRORS = 0 , 
    FIELDTERMINATOR = '','', 
    ROWTERMINATOR = ''\n''  
    )'
EXEC(@sql)

--Insert all from temp table to courses with extra column (user id)
INSERT INTO [courses]
SELECT @file_name as [user_id], [code], [subject], [type], [test], [grade]
FROM [temp_table]
GO

--Return every subject and its average 
CREATE PROC subjects_average
AS
SELECT [code],[subject], avg(convert(float, [grade]))
AS [Course Average]
FROM [courses] 
where isnumeric([grade])=1
GROUP BY [code],[subject]
GO

CREATE PROC user_subjects_average @user_id nvarchar(30)
AS
SELECT [subject], avg(convert(float, [grade]))
AS [Course Average]
FROM [courses] 
where isnumeric([grade])=1 AND [user_id] = @user_id
GROUP BY [subject]
GO

