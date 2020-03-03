# Ingrammicro SLA Checker and Email Monitoring Tool

![image](https://user-images.githubusercontent.com/45601866/75828876-d1635d00-5de7-11ea-998b-ffa14a1b8e4c.png)

<b>SLA Checker Tool</b> and <b>Email Monitoring Script</b> built in <b>Python</b>, is a script to monitor emails in a certain corporation and be able to keep track of server jobs and scripts and validate their SLA (Service Level Agreement) in an SSH Server, to check if a certain job when executed get passed on time or not.
The script is capable of generating a reports that can be easily understand. Below are the following formats of reports the script is capable of ..
<ul>
  <li><b>CSV Report</b> - Containing the summary of all jobs executed in the SSH Server including the time the job started to run and its completion, also with the SLA. It shows a <b>True or False</b> format in checking of Starting and Completion of time of a specific job. The report also shows the emails the system receive.</li>
  <li><b>TXT LOG Report Format</b> - Contains a hierarchy of all jobs executed in the SSH Server including the email receive, the respective job name of each emails, key sources from which each jobs are fall under, and the over all hierarchical structures of executions in the SSH Server.</li>
  <li><b>Email Report</b> - The script will going to send an Email to the corporate containing the reports of each jobs executed in the SSH Server. The email displays the <b>late jobs</b> executed in the server and the <b>missing jobs</b> that is currently not executed yet.</li>
</ul>

<b>Email Monitoring Report</b><br/>
Below is the sample of Email Report that the script will generate.<br/>
<b>Figure 1.1</b>

![image](https://user-images.githubusercontent.com/45601866/75426507-c6d83c00-597f-11ea-98b9-e62faec3c1d0.png)

<b>Figure 1.2</b>

![image](https://user-images.githubusercontent.com/45601866/75426621-06068d00-5980-11ea-9754-45f454284307.png)

<b>CSV Report</b><br/>
Below is also the type of report the script can generate in a <b>CSV</b> Format.<br/>
<b>Figure 1.3</b>

![dasdasdasd](https://user-images.githubusercontent.com/45601866/75427846-32bba400-5982-11ea-8b7c-48d2ced4f421.png)

<b>Job Hierarchy</b><br/>
Below is job hierarchy report the script can generate in a <b>TXT</b> format.<br/>
<b>Figure 1.4</b>

![ksadjkasldjkasd](https://user-images.githubusercontent.com/45601866/75500491-0f3b3c80-5a08-11ea-8324-3616d4971207.png)


