# Для запуска необходимо:
1) открыть в photoshop и исправить файл example.psd
2) создать .csv по примеру из example_studs.csv

### Дефолтная папка для результатов - results (параметр --results-dir)

### Дефолтный .csv - stud.csv (параметр --csv)

### Запрос для выгрузки csv
docker compose exec -it pg psql -U postgres -d Reports -c "COPY (SELECT id, course_n, group_n, surname, name, f_name, is_foreigner FROM reports_student) TO STDOUT WITH CSV HEADER" > stud_for_badges.csv