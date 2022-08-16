# import excel

从excel导入数据到数据库

## 配置

需要配置MYSQL服务的Host/用户名/密码，通过环境变了进行配置

```shell
export MYSQL_USER="root"
export MYSQL_PWD="xxxxxx"
export MYSQL_SERVER="10.0.0.1"
```

## 启动

```shell
go run main.go
```