<h2>Overview: <br /></h2>
&emsp;Working time registration system based on a Raspberry Pi Zero microcomputer and an RS522 RFID reader (13.56 MHz). The graphical interface is auto-generated .xslx files.

<h2>Gui:</h2>

![](img/GraphicInterface.png)

<h4> • Sheet "DANE" - Employees basic and payment information <br /></h4>
    <h4>Columns desc: </h4>
    &emsp;A - Rfid token id. <br />
    &emsp;B - Employee name. <br />
    &emsp;C - Salary per hour. <br />
    &emsp;D - Summary working time completed by the employee. <br />
    &emsp;E - Advanced payment. <br />
    &emsp;F - Insurance cost. <br />
    &emsp;G - Payment with advance payment and insurance deducted. <br />
    &emsp;H - Vacation days used. <br />
    &emsp;I - Employee status. Cell used to adding and deleting employees, more info in user manual. <br />
    &emsp;J2 - Auto-generated control value that represent amount of employees. DO NOT CHANGE THIS VALUE <br />
    &emsp;J4 - Last read unknown token id. <br />
    &emsp;L - Month calendar with daily employee working time, Weekends and holidays are marked in yellow. <br />
    
<h4> • Sheet "GODZINY" - detailed employees entrance info</h4>
  
<h2>User manual:</h2>

<h3> • Adding new employee:<br /></h3>
&emsp;To add new employee you need to fill A, B, C columnd and, mark "Z" on I column and reopen file. After reopen file new columns on calendar in "DANE" and "GODZINY" sheet should be added, and employee working status should change to "T".

<h3> • Removing employee:<br /></h3>
&emsp;To remove employee you need to change employee status to "N". This employee will be not included in the next month data sheet.

<h3> • Changing employee data<br /></h3>
&emsp;To change salary, vacation and advance payment simply change the value in the appropriate cell.

<h2> • Safe mode<br /></h2>
&emsp;When program detect error, system switch safe mode. When safe mode is on yellow led blink in 1 second period. Error information should be printed in <i>myapp.log</i>. If error containg wrong employee data, this data should be marked in red, additionally, a description of the error will be displayed under the employee data.
When employee hits the token, entrance time and employee token id will be saved in <i>myapp.log</i>.

<h3> • <br /></h3>

<h2>Leds and sound:</h2>
&emsp; 2 short buzzer beeps - token read correctly.
&emsp; long (2 second) buzzer beep - unknown tokenreaded, tokend id saved in J4 cell.
&emsp; Green led blink in 1 second period - System work correctly, ready to reading token.
&emsp; Yellow led blink in 1 second period - System work SafeMode.
&emsp; Yellow light, System is busy, not ready to read token. Yellow light shows for short time after reading token.
&emsp; No light or constantly glowing light - undefined behavior.

<h2>File access:</h2>
&emsp;To access the .xslx file on other computers I suggest to share program directory via samba server.

<h2>Electric diagram:</h2>

![](img/TimeCardsDiagram.jpg)
