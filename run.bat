echo "starting script"
cd C:\Project\UniLeverO2C
mvn clean test -Dcucumber.options="--tags @TestKato1"
echo "ending script"
pause