function sendEmails() {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const leadInstructorEmail = "dnewsom@codingdojo.com";
    const studentSupportEmail = "support@codingdojo.com";
    const doNotSendList = ['POSTPONED (PPD)', 'DROPPED (D)', 'ROLL BACK (RB)', 'LIMBO (L)', 'PAUSED (P)', 'EXPELLED (E)', 'MOVEDFORWARD (MF)', 'SELFPACED (SP)', 'TEXAS (T)', 'AUDITOR (SP)', 'DISMISSAL REVIEW (DR)'];
    const fullstackCourses = ['PYTHON', 'JAVA', 'MERN'];

    //Index Column Names
    const FIRST_NAME = 0
    const LAST_NAME = 1
    const EMAIL = 2
    const PROGRAM = 3
    const NEXT_STACK = 4
    const SSM = 5
    const CSM = 6
    const NOTES = 7
    const ROLLBACKS = 8
    const PROBATIONS = 9
    const EXAM_STATUS = 10
    const STATUS = 11
    const ASGMT_PER = 12
    const DQ_PER = 13
    const DQ_MISSED = 14

    const ssm_emails = {
        'Carleigh': 'cyoung@codingdojo.com',
        'Kit': 'dvirtudez@codingdojo.com',
        'Hannah/Robbie': 'hbouvier@codingdojo.com, rhannan@codingdojo.com',
    };

    const assignment_totals = {
        'WEBFUN': [2, 4, 6, 7],
        'PYTHON': [2, 4, 8, 12, 14, 16, 16, 16],
        'MERN': [3, 5, 6, 9, 10, 12, 13, 15],
        'JAVA': [2, 5, 8, 9, 10, 12, 13, 13],
        'PROJECTS AND ALGORITHMS': [3, 5, 7, 8]
    };

    let statusCounts = {
        'ON PACE (OP)': 0,
        'FALLING BEHIND (B)': 0,
        'AP WARNING (APW)': 0,
        'PROBATION (AP)': 0,
    };

    if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com" || Session.getActiveUser().getEmail() == "program-status@codingdojo.com") {
        //Get Class Info (Instructor Name, Email, Stack)
        var classInfo = ss.getRange(ss.createTextFinder('INSTRUCTOR DETAILS').findNext().getRow() + 1, 2, 4, 1).getValues();
        const instructor = classInfo[0][0];
        const instructorEmail = classInfo[2][0];
        const stack = classInfo[3][0];
        let emailCC = `${studentSupportEmail}, ${instructorEmail}`;

        //if Projects and Algos sheet offset columns by 3
        const offset = stack == 'PROJECTS AND ALGORITHMS' ? 3 : 0;

        const curr_week = parseInt(ss.getRange(2, 13 + offset).getValue()[5]);
        const duration = fullstackCourses.includes(stack) ? 8 : 4;

        //Get Assignment and Discussion column for current week
        let assignment_column = ss.getRange(2, 14 + offset, 1, ss.getLastColumn()).createTextFinder(`Week ${curr_week}`).findNext().getColumn() - 1
        let discussion_column = assignment_column + 1

        //If its an Odd Week, we'll be sending unofficial emails
        const week_type = curr_week % 2 === 1 ? 'Unofficial': 'Official'

        //Get all students whose status is not in the doNotSendList array
        let i = 4;
        let emailStudents = [];
        let discussionEmailCount = 0;
        while (ss.getRange(i, 1).getBackground() !== "#000000") {
            if (ss.getRange(i, 1).getValue() != '') {
                let status = ss.getRange(i, 12 + offset).getValue();
                let discussions_missed = parseInt(ss.getRange(i, 15 + offset).getValue());
                if (!doNotSendList.includes(status)) {
                    emailStudents.push(ss.getRange(i, 1, 1, ss.getLastColumn()).getValues());
                    statusCounts[status]++;
                    if (discussions_missed) {
                        discussionEmailCount++;
                    }
                }
            }
            i++
        }
        //Create Confirmation Popup
        let result = SpreadsheetApp.getUi().alert(`✉ Send ${week_type} Emails Week ${curr_week}`, `You're about to send ${week_type} emails to ${emailStudents.length} students including \n ${prettyPrint(statusCounts)}\n You will also be sending ${discussionEmailCount} discussion warning emails`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

        //If OK is clicked, start main loop
        if (result === SpreadsheetApp.getUi().Button.OK) {
            for (let student of emailStudents) {
                let curr_student = student[0];
                let email = curr_student[EMAIL];
                let status = curr_student[STATUS + offset];
                let missed_discussions = curr_student[DQ_MISSED + offset];
                let ssm_name = curr_student[SSM + offset];
                let ssm_email = ssm_emails[ssm_name];

                let message = generateMessage(curr_student, offset, assignment_column, discussion_column, stack, curr_week, ssm_name, ssm_email, instructor, week_type);

                //Add lead instructor emails to cc list for these 3 statuses
                if (status === 'FALLING BEHIND (B)' || status === 'AP WARNING (APW)' || status == 'PROBATION (AP)') {
                    emailCC += `, ${leadInstructorEmail}`;
                }

                // if Drew is Testing, don't cc anyone and send him all the emails
                if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com") {
                    emailCC = '';
                    email = 'dadorno@codingdojo.com';
                }

                //SEND EMAILS if email is a probation email, attach PDF
                if (status !== 'PROBATION (AP)' || week_type == 'Unofficial') {
                    MailApp.sendEmail(email, message[0], message[1], {
                        htmlBody: message[1],
                        cc: emailCC
                    });
                } else {
                    MailApp.sendEmail(
                        email,
                        message[0],
                        message[1], {
                            htmlBody: message[1],
                            cc: emailCC,
                            attachments: [DriveApp.getFileById('1-G1f-HbbxyFGkXDBq9joJ9zAkIzDmJKH').getAs(MimeType.PDF)]
                        }
                    );
                }
                //SEND MISSED DISCUSSION EMAILS if student has any missed Discussions
                if (missed_discussions) {
                    let discussionMessage = generateDiscussionMessage(curr_student, offset, instructor, stack, curr_week, duration, ssm_name, ssm_email)
                    emailCC = `${studentSupportEmail},${instructorEmail},${leadInstructorEmail}`
                    if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com") {
                        emailCC = '';
                    }
                    MailApp.sendEmail(email, discussionMessage[0], discussionMessage[1], {
                        htmlBody: discussionMessage[1],
                        cc: emailCC
                    })
                }
                //Only send Drew 1 test email
                if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com") {
                    break
                }
            }
            ssPrint(`${emailStudents.length + discussionEmailCount} emails were sent.`);
        } else {
            ssPrint("No emails were sent.");
        }
    } else {
        SpreadsheetApp.getUi().alert("You are not authorized to send emails.", SpreadsheetApp.getUi().ButtonSet.OK);
    }

    // Helper Functions*************************************************************************************************************************************************************

    //Creates the email template that will go out to each student
    function generateMessage(curr_student, offset, assignment_column, discussion_column, stack, curr_week, ssm_name, ssm_email, instructor, week_type) {
        let name = curr_student[FIRST_NAME];
        let assignment_completion = curr_student[ASGMT_PER + offset].toFixed(0);
        let discussion_completion = curr_student[DQ_PER + offset].toFixed(0);
        let status = curr_student[STATUS + offset];


        let completed_assignments = curr_student[assignment_column]
        let completed_discussions = curr_student[discussion_column];

        if(status == 'PROBATION (AP)' && week_type == 'Unofficial'){
            let one_on_one=curr_student[assignment_column-1]
            var one_on_one_status=one_on_one == "✅"?'✅':'❌'
            var assignment_status=assignment_completion >= 60 ?'✅':'❌'
                
            
        }

        let total_assignments = assignment_totals[stack][curr_week - 1]
        let total_discussions = curr_week * 2;

        // handling Hannah/Robbie students
        if (ssm_name === 'Hannah/Robbie') {
            ssm_name = 'Hannah or Robbie';
            ssm_email = '<a href="hbouvier@codingdojo.com">hbouvier@codingdojo.com</a> or <a href="rhannan@codingdojo.com">rhannan@codingdojo.com</a>';
        } else {
            ssm_email = `<a href="${ssm_email}">${ssm_email}</a>`;
        }

        let templateDict = {
            'ON PACE (OP)': [`Coding Dojo - ${stack} - Weekly Progress Update - Week ${curr_week}`,`<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the <b>${stack}</b> stack with<b> Instructor ${instructor}</b>. All students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions</p><p>Current Standing:</p><ul><li><b>Satisfactory Progress - Keep up the good work!</b></li><li>Assignments completion: ${assignment_completion}%</li><ul><li>${completed_assignments} of ${total_assignments} Core Assignments Completed</li></ul><li>Discussion completion: ${discussion_completion}%</li><ul><li>${completed_discussions} of ${total_discussions} Discussions Completed</li></ul><p>*Failure to maintain at minimum 90% completion of Core Assignments at the following progression check will result in being placed on an Academic Improvement Plan. Falling below 60% completion of Core Assignments will result in being placed on Academic Probation. Students who fail to meet Satisfactory Progress by the end of the stack will be placed on Academic Review to determine the course of action, which may include: retake of stack, withdrawal from program or program transfer (Web Fundamentals stack only).</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email}, for any questions or concerns.</p></ul>`],

            'FALLING BEHIND (B)': [`Coding Dojo - ${stack} - Academic Improvement Plan`,`<p>Hello ${name},</p><p>This notice serves to inform you that you have been placed on an <b>Academic Improvement Plan</b> due to your falling below 90% of Core assignment completion. This improvement plan is not meant as a punitive measure but to help you on your journey to academic recovery. </p><p>As a Student on an <b>Academic Improvement Plan</b>, it is recommended for you to complete the following:</p><ol><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a minimum of 1 Code Review per week with Instruction Staff.</li><li>Attend all scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <b>${stack}</b> stack with <b>Instructor ${instructor}</b> is the following:</p><ul><li><b>Marginal Progress - ${status} </b></li><li>Assignment progress: <b>${assignment_completion}%</b></li><ul><li>${completed_assignments} of ${total_assignments} Core Assignments Completed</li></ul><li>Discussion progress:<b>${discussion_completion}%</b></li><ul><li>${completed_discussions} of ${total_discussions} Discussions Completed</li></ul></ul><p>To Return to<b>Satisfactory Progress</b>, you must complete the following:</p><ol><li>Maintain a minimum of 90% Core assignment completion.</li><li>Maintain a minimum of 80% Discussion completion.</li></ol><p>Students who fail to meet all Academic Improvement Plan requirements and/or fail to meet progression standards will remain on an Academic Improvement Plan at the following progression check during the stack. Students who fail to reach satisfactory progress by the end of the stack will be placed under Academic Review to determine the course of action, which may include: Retake of Stack, Withdrawal from Program, Program Transfer (Web Fundamentals Stack only). A Student who falls below 60% of Core Assignment completion, as determined through regular review of student progress or meets other Academic Probation Criteria will be placed on Academic Probation. Three or more instances of unexcused academic probation on a student’s record will be grounds for academic dismissal from the program.</p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact your Student Support Manager, ${ssm_name} at ${ssm_email}</p>.`],

            'AP WARNING (APW)': [`Coding Dojo - ${stack} - Academic Improvement Plan`,`<p>Hello ${name},</p><p>This notice serves to inform you that you have been placed on an <b>Academic Improvement Plan</b> due to your falling below 90% of Core assignment completion. This improvement plan is not meant as a punitive measure but to help you on your journey to academic recovery. </p><p>As a Student on an <b>Academic Improvement Plan</b>, it is recommended for you to complete the following:</p><ol><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a minimum of 1 Code Review per week with Instruction Staff.</li><li>Attend all scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <b>${stack}</b> stack with <b>Instructor ${instructor}</b> is the following:</p><ul><li><b>Marginal Progress - ${status} </b></li><li>Assignment progress: <b>${assignment_completion}%</b></li><ul><li>${completed_assignments} of ${total_assignments} Core Assignments Completed</li></ul><li>Discussion progress:<b>${discussion_completion}%</b></li><ul><li>${completed_discussions} of ${total_discussions} Discussions Completed</li></ul></ul><p>To Return to<b>Satisfactory Progress</b>, you must complete the following:</p><ol><li>Maintain a minimum of 90% Core assignment completion.</li><li>Maintain a minimum of 80% Discussion completion.</li></ol><p>Students who fail to meet all Academic Improvement Plan requirements and/or fail to meet progression standards will remain on an Academic Improvement Plan at the following progression check during the stack. Students who fail to reach satisfactory progress by the end of the stack will be placed under Academic Review to determine the course of action, which may include: Retake of Stack, Withdrawal from Program, Program Transfer (Web Fundamentals Stack only). A Student who falls below 60% of Core Assignment completion, as determined through regular review of student progress or meets other Academic Probation Criteria will be placed on Academic Probation. Three or more instances of unexcused academic probation on a student’s record will be grounds for academic dismissal from the program.</p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact your Student Support Manager, ${ssm_name} at ${ssm_email}</p>.`],

            'PROBATION (AP)': [`Coding Dojo - ${stack} - Academic Probation`,`<p>Hello ${name},</p><p>This notice is to inform you that you have been placed on Academic Probation due to failing to meet the academic progress requirements as outlined in the catalog and having made unsatisfactory progress. Coding Dojo places students on Academic Probation with less than 60% Core Assignment submission. We would like to see you return to Satisfactory Progress, and we support you as you take the steps to academic recovery.</p><p>As a Student on <b>Academic Probation</b>, you are required to complete the following:</p><ol><li>Contact your Student Support Manager and current instructor to schedule regular updates on their academic progress.</li><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a <b>minimum</b> of 1 Code Review per week with Instruction Staff.</li><li>Attend <b>all</b> scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <b>${stack}</b> stack with<b>Instructor ${instructor}</b> is the following:</p><ul><li><b>Academic Probation</b></li><li>Assignment progress:<b>${assignment_completion}%</b><ul><li>${completed_assignments} of ${total_assignments} Assignments Completed</li></ul></li><li>Discussion progress: ${discussion_completion}%</li><ul><li>${completed_discussions} of ${total_discussions} Discussions Completed</li></ul></ul><p>To return to Marginal Progress and no longer be on Academic Probation, you must complete the following:</p><ol><li>Fulfill the requirements listed above.</li><li>Maintain a minimum of <b>60%</b> Core assignment completion.</li><li>Maintain a minimum of <b>80%</b> discussion completion.</li></ol><p>Students who meet all Academic Probation requirements and/or fail to meet progression standards will be placed on an additional instance of Academic Probation at the following progression check during the stack. Students who fail to reach good standing by the end of the stack will be placed under Academic Review to determine the course of action, which may include: Retake of Stack, Withdrawal from Program, Program Transfer (Web Fundamentals Stack only). Three or more instances of unexcused academic probation on a student’s record will be grounds for academic dismissal from the program.</p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact your Student Support Manager, ${ssm_name} at ${ssm_email}</p>. A student’s status of Academic Probation may be lifted once the student returns to Marginal Progress standing (defined above) at minimum or is excused. A student is allowed one excused Academic Probation per stack</p>`],
        }

        let unofficialTemplateDict={
            'ON PACE (OP)': [`Coding Dojo - ${stack} - Unofficial Progress Update`,`<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the <strong>${stack}</strong> stack with Instructor <strong>${instructor}</strong>. Official progress check-ins will be sent out next week that will determine your academic standing within the program. This check-in is meant as a pulse update so that yourself, Instruction and Support can have an update on where you're at percentage wise.</p><p>As of this week your Current standing is:</p><ul><li aria-level="1"><strong>On Pace - Keep up the good work!</strong></li></ul><ul><li aria-level="1">Assignments completion: <strong>${assignment_completion}</strong></li><ul><li aria-level="2"><strong>${completed_assignments} of ${total_assignments} Core Assignments Completed</strong></li></ul></ul><p>As a reminder, all students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions.</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} or the Student Support team at support@codingdojo.com, for any questions or concerns.</p>`],

            'FALLING BEHIND (B)':[`Coding Dojo - ${stack} - Unofficial Progress Update`,`<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the <strong>${stack}</strong> stack with Instructor <strong>${instructor}</strong>. Official progress check-ins will be sent out next week that will determine your academic standing within the program. This check-in is meant as a pulse update so that yourself, Instruction and Support can have an update on where you're at percentage wise.</p><p>As of this week your Current standing is:</p><ul><li aria-level="1"><strong>Falling Behind</strong></li></ul><ul><li aria-level="1">Assignments completion: <strong>${assignment_completion}</strong><strong>}</strong></li><ul><li aria-level="2"><strong>${completed_assignments} of ${total_assignments} Core Assignments Completed</strong></li></ul></ul><p>Falling Behind means that you are currently within the range of 60%-89% completion of Core Assignments. This would be considered Marginal Progress at the time of an official check-in and if progress remains within the same range you would be placed on an Academic Improvement Plan, which is not a punitive status and is simply meant to help students be successful within the program.</p><p>If you feel that you are in need of additional instructional support be sure to reach out to your Instructor, TAs in Discord or to schedule a Code Review. Should you feel like you are having trouble with keeping up with assignments please contact your Student Support Manager.</p><p>As a reminder, all students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions.</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} or the Student Support team at support@codingdojo.com, for any questions or concerns.</p>`],

            'AP WARNING (APW)':[`Coding Dojo - ${stack} - Unofficial Progress Update`,`<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the <strong>${stack}</strong> stack with Instructor <strong>${instructor}</strong>. Official progress check-ins will be sent out next week that will determine your academic standing within the program. This check-in is meant as a pulse update so that yourself, Instruction and Support can have an update on where you're at percentage wise.</p><p>As of this week your Current standing is:</p><ul><li aria-level="1"><strong>Academic Probation Warning&nbsp;</strong></li></ul><ul><li aria-level="1">Assignments completion: <strong>${assignment_completion}</strong></li><ul><li aria-level="2"><strong>${completed_assignments} of ${total_assignments} Core Assignments Completed</strong></li></ul></ul><p>Academic Probation Warning means that you are currently within the range of 0%-59% completion of Core Assignments. This would be considered Unsatisfactory Progress at the time of an official check-in and if progress remains within the same range you would be placed on Academic Probation, which is a punitive status and a student who receives more than 2 official notices of Academic Probation will be recommended for Academic Dismissal.</p><p>As a reminder, all students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions.</p><p>If you feel that you are in need of additional instructional support be sure to reach out to your Instructor, TAs in Discord or to schedule a Code Review. Should you feel you are having trouble keeping up with assignments, or something has come up that affects your ability to progress with the program please contact your Student Support Manager.</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} or the Student Support team at support@codingdojo.com, for any questions or concerns.</p>`],

            'PROBATION (AP)':[`Coding Dojo - ${stack} - Unofficial Progress Update - Outstanding AP Requirements Not Completed`,`<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the<strong> ${stack} </strong>stack with Instructor<strong> ${instructor} </strong>. Official progress check-ins will be sent out next week that will determine your academic standing within the program. This check-in is meant as a pulse update so that yourself, Instruction and Support can have an update on where you're at percentage wise.</p><p>As of this week your Current standing is:</p><ul><li aria-level="1"><strong>Academic Probation</strong></li></ul><ul><li aria-level="1">Assignments completion:<strong>${assignment_completion}</strong></li><ul><li aria-level="2"><strong>${completed_assignments} of ${total_assignments} Core Assignments Completed</strong></li></ul></ul><p><strong>**This update signifies that you were placed on Academic Probation at the last official check-in for your current stack and that you are not yet meeting the requirements to be taken off of Academic Probation.</strong></p><p>Requirements include:</p><ul style="list-style:none"><li aria-level="1"><strong>${assignment_status}&ensp; Have at least 60% Core Assignment completion</strong></li><li aria-level="1"><strong>${one_on_one_status}&ensp; Have met with your Instructor since being placed on Academic Probation</strong></li></ul><p><strong>A student who receives more than 2 official notices of Academic Probation will be recommended for Academic Dismissal.</strong></p><p>As a reminder, all students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions.</p><p>If you feel that you are in need of additional instructional support be sure to reach out to your Instructor, TAs in Discord or to schedule a Code Review. Should you feel like you are having trouble keeping up with assignments, or something has come up to affect your ability to progress with the program please contact your Student Support Manager</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} or the Student Support team at support@codingdojo.com, for any questions or concerns.</p>`]
        }

        return week_type=='Official' ? templateDict[status] : unofficialTemplateDict[status]
    }
    //Creates the missed discussion email template that will go out to students that are missing discussions
    function generateDiscussionMessage(curr_student, offset, instructor, stack, curr_week, duration, ssm_name, ssm_email) {
        let name = curr_student[FIRST_NAME];
        let discussion_completion = curr_student[DQ_PER + offset].toFixed(0);
        let missed_discussions = curr_student[DQ_MISSED + offset];
        let total_discussions = curr_week * 2;
        let completed_discussions = total_discussions - missed_discussions;

        if (ssm_name === 'Hannah/Robbie') {
            ssm_name = 'Hannah or Robbie';
            ssm_email = '<a href="hbouvier@codingdojo.com">hbouvier@codingdojo.com</a> or <a href="rhannan@codingdojo.com">rhannan@codingdojo.com</a>';
        } else {
            ssm_email = `<a href="${ssm_email}">${ssm_email}</a>`;
        }

        let templateDict = {
            'behind': ['Coding Dojo - Discussion Warning', `<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week. You may already be in contact with your instructor or student support regarding this matter.</p><p>This is a warning regarding your <strong>Discussion Completion</strong> during the <strong>${stack} </strong>stack with Instructor <strong>${instructor}</strong>. All students are expected to maintain above 80% in Discussion Completion and above 90% completion of Core Assignments.</p><p>Current Discussion Standing: <strong>Discussion Warning</strong></p><ul><li>Discussion completion: <strong>${discussion_completion}</strong></li><ul><li><strong>${completed_discussions} of ${total_discussions} Discussions Completed</strong></li></ul><li>Discussions Missed: <strong>${missed_discussions}</strong></li></ul><p>A student who has not logged on to the Online Learning Platform and posted on their assigned Discussion for more than five (5) consecutive assigned sessions, is considered inactive. In this instance, the student will face termination from the bootcamp if they are unable to be reached by Coding Dojo staff. A student who has missed more than 20% of the required discussions by the end of the stack will be required to withdraw from the program.</p><p>Campus staff may excuse up to 10% of a student's missed discussions for special or mitigating circumstances outside the control of the student. In those cases, the circumstances must be provided, in writing, to campus staff as soon as possible.</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} if you have any questions.</p>`],

            'final': ['Coding Dojo - Discussion FINAL Warning', `<p>Hello ${name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week. You may already be in contact with your instructor or student support regarding this matter.</p><p>This is a warning regarding your <strong>Discussion Completion</strong> during the <strong>${stack} </strong>stack with Instructor <strong>${instructor}</strong>. All students are expected to maintain above 80% in Discussion Completion and above 90% completion of Core Assignments.</p><p>Current Discussion Standing: <strong>Discussion FINAL Warning</strong></p><ul><li>Discussion completion: <strong>${discussion_completion}</strong></li><ul><li><strong>${completed_discussions} of ${total_discussions} Discussions Completed</strong></li></ul><li>Discussions Missed: <strong>${missed_discussions}</strong></li></ul><p>A student who has not logged on to the Online Learning Platform and posted on their assigned Discussion for more than five (5) consecutive assigned sessions, is considered inactive. In this instance, the student will face termination from the bootcamp if they are unable to be reached by Coding Dojo staff. A student who has missed more than 20% of the required discussions by the end of the stack will be required to withdraw from the program.</p><p>Campus staff may excuse up to 10% of a student&rsquo;s missed discussions for special or mitigating circumstances outside the control of the student. In those cases, the circumstances must be provided, in writing, to campus staff as soon as possible.</p><p>Feel free to reach out to your Student Support Manager, ${ssm_name} at ${ssm_email} if you have any questions.</p>`]
        };
        return missed_discussions < duration / 2 ? templateDict['behind'] : templateDict['final'];
    }

    function prettyPrint(obj) {
        returnStr = ''
        for (let [key, value] of Object.entries(obj)) {
            returnStr += `${key.substring(0, key.indexOf("("))}: ${value}\n`;
        }
        return returnStr;
    }

    function ssPrint(str) {
        SpreadsheetApp.getActive().toast(str);
    }
}