# PEI Firm Finder

Java + Apache POI app that reads a PEI 300 sheet, filters by Region/AUM/Latest Fund Size/Focus, ranks by a user-defined priority, and writes `PEI300_SortedFile.xlsx`.

## Requirements
- Java 17+
- Maven

## Run
```bash
mvn -q -DskipTests package
java -cp target/$(ls target | grep .jar$ | head -n1) poi_example.Demo
