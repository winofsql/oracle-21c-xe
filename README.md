# [oracle-21c-xe ( setup.exe を実行 )](https://www.oracle.com/jp/database/technologies/xe-downloads.html)

![image](https://user-images.githubusercontent.com/1501327/174948000-a071ae2f-03da-438f-9b7e-1496a74ac65d.png)

![image](https://user-images.githubusercontent.com/1501327/174948103-5cccd7d8-7ccb-4dd8-a9ce-a47708085a03.png)

![image](https://user-images.githubusercontent.com/1501327/174948165-55195fb8-8141-4fba-8981-1de55272333b.png)

![image](https://user-images.githubusercontent.com/1501327/174948216-9d694b5b-e3ae-4977-b7ac-9465beffe33e.png)

![image](https://user-images.githubusercontent.com/1501327/174950917-bb429b80-8be2-4d0a-8a38-a55b61f5ceca.png)

![image](https://user-images.githubusercontent.com/1501327/174949924-8bd60759-06d8-45ce-a25c-df1530a89b2f.png)

![image](https://user-images.githubusercontent.com/1501327/174952223-ca78d178-ef1c-432b-9a90-2c12d5494dfc.png)

![image](https://user-images.githubusercontent.com/1501327/174952300-c6b1f9b9-5be0-4317-808b-5f9b3abb5a25.png)

![image](https://user-images.githubusercontent.com/1501327/174952530-966ec3e9-80fa-468d-b609-9400a3dba0c9.png)

### "C:\app\lightbox\product\21c\homes\OraDB21Home1\network\admin\tnsnames.ora"
```
XEPDB1 =
  (DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = R101-00)(PORT = 1521))
    (CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = XEPDB1)
    )
  )
```

### NLS_LANG : Japanese_Japan.JA16SJIS

![image](https://user-images.githubusercontent.com/1501327/174953229-036e9154-420f-4046-afad-79fe7dd0f6d0.png)

[Oracle Instant Client Downloads](https://www.oracle.com/database/technologies/instant-client/downloads.html)\
![image](https://user-images.githubusercontent.com/1501327/186780032-1fee9b62-1acc-4797-ab5c-363c6fafc3f1.png)

![image](https://user-images.githubusercontent.com/1501327/186783062-bc93a95c-0a87-4beb-9ec8-7e4ee7569403.png)

![image](https://user-images.githubusercontent.com/1501327/186783145-c54c79c7-69da-467e-8314-159cb0bce84b.png)

![image](https://user-images.githubusercontent.com/1501327/174955274-6d628f6f-4221-43fd-b251-b37aa288836e.png)

![image](https://user-images.githubusercontent.com/1501327/174955605-3e47802c-61b8-4db0-9bf5-3500151abb76.png)

```sql
create tablespace LIGHTBOX00_SPACE
	datafile 'C:\app\lightbox\product\21c\oradata\XE\LIGHTBOX00PDB.DBF'
	size 5M
	autoextend on
	next 1M
	maxsize unlimited
	segment space management AUTO;

create user LIGHTBOX00
	identified by trustno1
	default tablespace LIGHTBOX00_SPACE
	temporary tablespace TEMP
	quota unlimited on LIGHTBOX00_SPACE
	account UNLOCK;
	
grant 
	 ALTER PROFILE
	,ALTER SESSION
	,ALTER SYSTEM
	,ALTER TABLESPACE
	,ALTER USER
	,CREATE ANY DIRECTORY
	,CREATE PROCEDURE
	,CREATE PROFILE
	,CREATE PUBLIC SYNONYM
	,CREATE ROLE
	,CREATE ROLLBACK SEGMENT
	,CREATE SEQUENCE
	,CREATE SESSION
	,CREATE SYNONYM
	,CREATE TABLE
	,CREATE TABLESPACE
	,CREATE TRIGGER
	,CREATE VIEW
	,DROP ANY DIRECTORY
	,EXECUTE ANY PROCEDURE
	,SELECT ANY DICTIONARY
	,SELECT ANY SEQUENCE
	,SELECT ANY TABLE
to LIGHTBOX00
```

![image](https://user-images.githubusercontent.com/1501327/174956335-ea53d665-8eb5-4ca0-9bf3-fca387d0477a.png)

### ファイアーウォール : 受信 : ポートの追加( 1521 )
![image](https://user-images.githubusercontent.com/1501327/174964628-bf818aeb-ff14-49b3-975f-db649d96d54c.png)


### データ作成
```
"C:\Windows\SysWOW64\cscript.exe" hanbaiC_oracle21c_nvarchar.vbs
```

### VSCode 拡張で接続
![image](https://user-images.githubusercontent.com/1501327/175455867-ec259ec5-3ba7-4cb1-8779-c464f9f7c0f9.png)\
![image](https://user-images.githubusercontent.com/1501327/175455814-eb964a8a-4f44-46ed-ac59-3ef5ab9c2b6c.png)\
![image](https://user-images.githubusercontent.com/1501327/175455721-e2b70bb3-e510-45e3-abef-61a6d87ca282.png)

![image](https://user-images.githubusercontent.com/1501327/175455953-0cbd6426-1b2a-46ae-8a4d-e635b9c44706.png)
