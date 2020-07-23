USE Database Name;

/*BASIC CLEAN UP - DROP BLANK ROWS, REFORMAT variable12*/

IF OBJECT_ID('???.TableName', 'U') IS NOT NULL DROP TABLE [???].[TableName2];
IF OBJECT_ID('???.TableName', 'U') IS NOT NULL SELECT * INTO [???].[TableName2] FROM ???.TableName;

UPDATE [???].[TableName2]
SET    variable10 = CASE WHEN variable10 = '11' THEN '1'
			WHEN variable10 = '22' THEN '2'
			WHEN variable10 = '33' THEN '3'
			WHEN variable10 = '44' THEN '4'
			WHEN variable10 = '55' THEN '5'
			WHEN variable10 = '66' THEN '6'
			WHEN variable10 = '77' THEN '7' 
			ELSE variable10 END;

UPDATE [???].[TableName2] 
SET variable9 = '??'
WHERE variable9 in ('Y', 'X') and variable10 in ('1', '2', '3', '4', '5');

DELETE FROM [???].[TableName2] 
WHERE variable154 in ('aaa', 'bbb', 'ccc');

DELETE FROM [???].[TableName2] 
WHERE variable9 in ('??', '??') and variable10 = '0' and variable12 = '0';

DELETE FROM [???].[TableName2] 
WHERE variable10 = '0' and variable14 = 0 and variable20 = 0 and variable90 = 0;

/* CLEAN UP variable12s */

UPDATE [???].[TableName2] 
SET variable12 = 'a'
WHERE variable12 = 'b';

UPDATE [???].[TableName2] 
SET variable12 = 'c'
WHERE variable12 = 'd';

UPDATE [???].[TableName2] 
SET variable12 = 'e'
WHERE variable12 = 'f';

/* Update variable10 */

UPDATE [???].[TableName2] 
SET variable10 = 'g', variable11 = 'h'
WHERE variable10 ='i';

UPDATE [???].[TableName2] 
SET variable10 = 'j', variable11 = 'k'
WHERE variable10 = 'l';

UPDATE [???].[TableName2] 
SET variable10 = 'm', variable11 = 'n'
WHERE variable10 = 'o' and variable11 = 'p';

/* Update variable9 */

UPDATE [???].[TableName2] 
SET variable9 = 'q'
WHERE variable10 = 'r';

UPDATE [???].[TableName2] 
SET variable9 = 's'
WHERE variable10 in ('t', 'u', 'v');

UPDATE [???].[TableName2] 
SET variable9 = 'w'
WHERE variable10 in ('x', 'y');

/* Rebuild variable1 Fields */

UPDATE [???].[TableName2]
SET variable3 = [variable8];

UPDATE [???].[TableName2]
SET variable4 = CASE WHEN [variable8] < variable153 then 'Ok'
			ELSE [variable8] END;

UPDATE [???].[TableName2]
SET variable5 = variable1(variable12, '\', [variable153]);

UPDATE [???].[TableName2]
SET variable6 = variable1(variable10,variable11,'\',[variable153]);

UPDATE [???].[TableName2]
SET variable7 = variable1([variable8],'\',variable9, '\', variable10,'\',variable11,'\',variable12);


/* Rebuild Fields That Didn't Exist Prior to Certain Files*/

UPDATE [???].[TableName2]
SET variable155 = CASE WHEN [variable] = variable153 then 0
                ELSE variable37 + variable63 - variable59 END;

UPDATE [???].[TableName2]
SET variable156 = CASE WHEN [variable] = variable153 then 0
				ELSE variable38 + variable64 - variable60 END