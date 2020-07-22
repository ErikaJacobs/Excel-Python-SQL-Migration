USE Database Name;

/* DELETE TABLE BEFORE RECREATING */

IF OBJECT_ID('???.table1', 'U') IS NOT NULL DROP TABLE [???].[table1];
IF OBJECT_ID('???.table2', 'U') IS NOT NULL DROP TABLE [???].[table2];

/* Final View - Numerical Values */

WITH a as (
		SELECT variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12, SUM(variable13) as "variable13", SUM(variable14) as "variable14", SUM(variable19) as "variable19", SUM(variable20) as "variable20", 
		(SUM(variable37)+SUM(variable63)-SUM(variable59)) as 'Lvariable99S', 
		(SUM(variable38) + SUM(variable64) - SUM(variable60)) as 'Desc',
		SUM(variable82) as 'variable80', 
		(SUM(variable38)+SUM(variable64)-SUM(variable60)+SUM(variable82)) as 'Desc',
		(SUM(variable37)+SUM(variable63)-SUM(variable59)+SUM(variable38)+SUM(variable64)-SUM(variable60)+SUM(variable82)) as 'Name2',
		(SUM(variable23)+SUM(variable76)+SUM(variable79)+SUM(variable77)+SUM(variable78)) as "name1"
		FROM [???].[TableName2]
		WHERE [variable8] = variable153
		GROUP BY variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12),
	b as (
		SELECT variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12, (SUM(variable37)+SUM(variable63)-SUM(variable59)) as 'variable100_Lvariable99S', 
		(SUM(variable38) + SUM(variable64) - SUM(variable60)) as 'variable100_Desc'
		FROM [???].[TableName2]
		GROUP BY variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12),
	c as (SELECT variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12, SUM(variable15) as "variable15", SUM(variable16) as "variable16", SUM(variable21) as "variable21", SUM(variable22) as "variable22",
		(SUM(variable24)+SUM(variable88)+SUM(variable86)+SUM(variable87)+SUM(variable85)) as "name1_Desc"
		FROM [???].[TableName2]
		WHERE [variable8] <> variable153
		GROUP BY variable153, variable152, variable154, variable7, [variable8], variable9, variable10, variable11, variable12),
d as (SELECT b.variable153, b.variable152, b.variable154, b.variable7, b.[variable8], b.variable9, b.variable10, b.variable11, b.variable12, 
ISNULL(a.variable13, 0) as 'variable13', ISNULL(a.variable14, 0) as 'variable14', ISNULL(a.variable19, 0) as 'variable19', ISNULL(a.variable20, 0) as 'variable20', 
(ISNULL(a.variable13,0)+ISNULL(c.variable15,0)) as 'variable100_variable13', 
(ISNULL(a.variable14, 0)+ISNULL(c.variable16,0)) as 'variable100_variable14', 
(ISNULL(a.variable19,0)+ISNULL(c.variable21,0)) as 'variable100_variable19', 
(ISNULL(a.variable20,0)+ISNULL(c.variable22,0)) as 'variable100_variable20',
ISNULL(a.Lvariable99S,0) as 'Lvariable99S', 
ISNULL(a.Desc,0) as 'Desc', 
ISNULL(a.variable80,0) as 'variable80', 
ISNULL(a.Desc,0) as 'Desc', 
ISNULL(a.name2,0) as 'name2',
b.variable100_Lvariable99S, b.variable100_Desc,
(b.variable100_Desc+ISNULL(a.variable80,0)) as 'variable100_Desc', 
(b.variable100_Lvariable99S+b.variable100_Desc+ISNULL(a.variable80,0)) as 'variable100_Desc',
(b.variable100_Lvariable99S - ISNULL(a.Lvariable99S,0)) as 'Lvariable99S',
(b.variable100_Desc - ISNULL(a.Desc,0)) as 'Desc',
(b.variable100_Lvariable99S - ISNULL(a.Lvariable99S,0)+b.variable100_Desc - ISNULL(a.Desc,0)) as 'Total',
ISNULL(a.name1,0) as 'name1', 
ISNULL(c.name1_Desc,0)+ISNULL(a.name1,0) as 'variable100_name1'
FROM b
FULL JOIN a on b.variable154 = a.variable154 and b.variable7 = a.variable7
FULL JOIN c on b.variable154 = c.variable154 and b.variable7 = c.variable7
),

/*Tables a, b, c, and d above are variable100 Metrics - THEY SHOULD NEVER CHANGE UNLESS ADDING METRICS */

/* To Change Groupings: Adjust Case Statements ONLY below - Once in Select, Once in Group By (Twice Total)*/
e as (
SELECT variable154, variable153, variable152, 'Group' = 
(CASE
	WHEN variable9 = 'C' THEN 'D'
	WHEN variable9 = 'A' THEN 'B'
	WHEN variable9 = 'M' THEN 'N'
	WHEN variable9 in ('W', 'X', 'Y') then 'Z'
	ELSE '999' END
END),
SUM(variable13) as 'variable13', SUM(variable14) as 'variable14', SUM(variable19) as 'variable19', SUM(variable20) as 'variable20',
SUM(variable100_variable13) as 'variable100_variable13', SUM(variable100_variable14) as 'variable100_variable14', SUM(variable100_variable19) as 'variable100_variable19', SUM(variable100_variable20) as 'variable100_variable20',
SUM(Lvariable99S) as 'name', SUM(Desc) as 'Desc', SUM(variable80) as 'variable80', SUM(Desc) as 'Desc',
SUM(name2) as 'name2', SUM(variable100_Lvariable99S) as 'variable100_Lvariable99S', SUM(variable100_Desc) as 'variable100_Desc',
SUM(variable100_Desc) as 'variable100_Desc', SUM(variable100_name2) as 'variable100_name2', SUM(name) as 'name', 
SUM(Desc) as 'Desc', SUM(Total) as 'Total', SUM(name1) as 'name1',
SUM(variable100_name1) as 'variable100_name1'
FROM d
GROUP BY variable154, variable153, variable152, (
CASE
	WHEN variable9 = 'C' THEN 'D'
	WHEN variable9 = 'A' THEN 'B'
	WHEN variable9 = 'M' THEN 'N'
	WHEN variable9 in ('W', 'X', 'Y') then 'Z'
	ELSE '999' END)
UNION
SELECT variable154, variable153, variable152, 'Group' = 
(CASE
	WHEN variable10 in ('A', 'B', 'C', 'D', 'E') or (variable10 = 'F' and variable11 = 'G') THEN 'H'
	WHEN variable10 in ('I', 'J', 'K', 'L') THEN 'M'
	WHEN variable10 in ('N', 'O', 'P', 'Q') THEN 'R'
	WHEN variable10 in ('S') THEN 'T'
	WHEN variable10 in ('U', 'V', 'W', 'X') or (variable10 = 'Y' and variable11 = 'Z') THEN 'AA'
	ELSE '999'
END),
SUM(variable13) as 'variable13', SUM(variable14) as 'variable14', SUM(variable19) as 'variable19', SUM(variable20) as 'variable20',
SUM(variable100_variable13) as 'variable100_variable13', SUM(variable100_variable14) as 'variable100_variable14', SUM(variable100_variable19) as 'variable100_variable19', SUM(variable100_variable20) as 'variable100_variable20',
SUM(Lvariable99S) as 'name', SUM(Desc) as 'Desc', SUM(variable80) as 'variable80', SUM(Desc) as 'Desc',
SUM(name2) as 'name2', SUM(variable100_Lvariable99S) as 'variable100_Lvariable99S', SUM(variable100_Desc) as 'variable100_Desc',
SUM(variable100_Desc) as 'variable100_Desc', SUM(variable100_name2) as 'variable100_name2', SUM(name) as 'name', 
SUM(Desc) as 'Desc', SUM(Total) as 'Total', SUM(name1) as 'name1',
SUM(variable100_name1) as 'variable100_name1'
FROM d
GROUP BY variable154, variable153, variable152, (
CASE
	WHEN variable10 in ('A', 'B', 'C', 'D', 'E') or (variable10 = 'F' and variable11 = 'G') THEN 'H'
	WHEN variable10 in ('I', 'J', 'K', 'L') THEN 'M'
	WHEN variable10 in ('N', 'O', 'P', 'Q') THEN 'R'
	WHEN variable10 in ('S') THEN 'T'
	WHEN variable10 in ('U', 'V', 'W', 'X') or (variable10 = 'Y' and variable11 = 'Z') THEN 'AA'
	ELSE '999' END)
	)

SELECT *
INTO [???].[table1]
From e
WHERE e.[Group] <> '999'
ORDER BY variable153 DESC, variable152 DESC, [Group];
;

/* Final View - Ratios */

WITH f as (
SELECT variable154, variable153, variable152, [Group],
ISNULL((name/NULLIF(variable100_variable20,0)),0) as 'name', 
ISNULL((Desc/NULLIF(variable100_variable20,0)),0) as 'Desc', 
ISNULL((variable80/NULLIF(variable100_variable20,0)),0) as 'variable80',
ISNULL((Desc/NULLIF(variable100_variable20,0)),0) as 'Desc', 
ISNULL((name+Desc+variable80)/NULLIF(variable100_variable20,0),0) as 'name2',
ISNULL((variable100_name/NULLIF(variable100_variable20,0)),0) as 'variable100_name', 
ISNULL((variable100_Desc/NULLIF(variable100_variable20,0)),0) as 'variable100_Desc', 
ISNULL((variable100_Desc/NULLIF(variable100_variable20,0)),0) as 'variable100_Desc', 
ISNULL((variable100_name+variable100_Desc+variable80)/NULLIF(variable100_variable20,0),0) as 'variable100_name2',
ISNULL((name/NULLIF(variable100_variable20,0)),0) as 'name', 
ISNULL((Desc/NULLIF(variable100_variable20,0)),0) as 'Desc', 
ISNULL((Total/NULLIF(variable100_variable20,0)),0) as 'Total',
ISNULL((variable100_name1/NULLIF(variable100_variable19,0)),0) as 'name1',
ISNULL((name+Desc+variable80)/NULLIF(variable100_variable20,0),0)+ISNULL((variable100_name1/NULLIF(variable100_variable19,0)),0) as 'name3',
ISNULL((variable100_name+variable100_Desc+variable80)/NULLIF(variable100_variable20,0),0)+ISNULL((variable100_name1/NULLIF(variable100_variable19,0)),0) as 'variable100_name3'
FROM [???].[table1]
)

SELECT *
INTO [???].[table2]
FROM f
ORDER BY variable153, variable152, [Group];