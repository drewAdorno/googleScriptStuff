function sendProgressUpdates() {
    let template = `<p>Hello {name},</p><p>This is a weekly update on your assignment and discussion completion</p><p>Core Assignment Completion: <b>{assignment_completion}%</b></p><p>Discussion Completion: <b>{discussion_completion}%</b></p><p>Exam Status: <b>{exam_status}</b></p><p>Remember to pass {stack}, you must have <b>80%</b> of your discussions and <b>90%</b> of your core assignments in and have passed the exam by the <b>{end_date}</b></p><p>Feel free to schedule some time with me at {instructor_calendly} to go over any questions or concerns you may have</p>`
    const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let classInfo = ss.getRange(ss.createTextFinder('INSTRUCTOR DETAILS').findNext().getRow() + 1, 2, 4, 1).getValues();

    const studentSupportEmail = "proehl@codingdojo.com";
    const doNotSendList = ['POSTPONED(PPD)', 'DROPPED(D)', 'ROLLBACK(RB)', 'LIMBO(L)', 'EXPELLED(E)', 'TRANSFERRED(T)'];
    const instructorName = classInfo[0][0];
    const instructorEmail = classInfo[2][0];
    const stack = classInfo[3][0];
    let end_date = Utilities.formatDate(ss.getRange(1, 2).getValue(), "EST", "MM/dd/yyyy")


    const calendlylinks = {
        'Drew': 'https://calendly.com/drew-adorno/one-on-one',
        'Michael': 'https://calendly.com/mclinkscales',
        'Chris': 'https://calendly.com/cjuarez-52'
    }

    let i = 4
    let emailStudents = [];
    let instructorCalendly = calendlylinks[instructorName]

    while (ss.getRange(i, 1).getBackground() !== "#000000") {
        if (ss.getRange(i, 1).getValue() != '') {
            let status = ss.getRange(i, 9).getValue();
            status = status.split(" ").join("");
            if (!doNotSendList.includes(status)) {
                emailStudents.push(ss.getRange(i, 1, 1, 11).getValues());
            }
        }
        i++
    }

    let result = SpreadsheetApp.getUi().alert(`You're about to send emails to ${emailStudents.length} students.`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (result === SpreadsheetApp.getUi().Button.OK) {
        for (let student of emailStudents) {
            let curr_student = student[0]
            let name = curr_student[0];
            let email = curr_student[2];
            let exam_status = curr_student[7]
            let status = curr_student[8];
            let assignment_completion = curr_student[9].toFixed(0);
            let discussion_completion = curr_student[10].toFixed(0);

            Logger.log({
                'name': name,
                'email': email,
                'assignment_completion': assignment_completion,
                'disscussion_completion': discussion_completion,
                'status': status,
                'exam status': exam_status
            })


            var message = template.replace("{name}", name).replace("{assignment_completion}", assignment_completion).replace("{discussion_completion}", discussion_completion).replace("{stack}", stack).replace("{instructor}", instructorName).replace("{end_date}", end_date).replace("{instructor_calendly}", instructorCalendly).replace("{exam_status}", exam_status).replace("{stack}", stack);

            MailApp.sendEmail(email, `${stack} - Weekly Progress Report`, message, {
                cc: `${studentSupportEmail}, ${instructorEmail}`,
                htmlBody: message
            });
        }
        ssPrint(`${emailStudents.length} emails were sent.`)
    } else {
        ssPrint("No emails were sent.");
    }
}

function sendEmails() {
    let old_templateDiscussionWarning = `<p>Hello {name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week. You may already be in contact with your instructor or student support regarding this matter.</p><p>This is a warning regarding your progress during the <strong>{stack}</strong> stack with <strong>Instructor {instructor}</strong>. All students are to maintain above 80% in discussion completion and above 70% completion of mandatory assignments.</p><p>Current Standing: <strong>Discussion Warning</strong></p><ul><li>Discussion completion: <strong> {discussion_completion}%</strong></li><li>Assignments completion: {assignment_completion}%</li></ul></p><p> Failure to maintain the required percentage stated above throughout the rest of the stack may result in being placed on an Academic Improvement Plan</p><p>Feel free to reach out to your {student_support}, for any questions or concerns.</p>`;

    let old_templateAssignmentWarning = "<p>Hello {name},</p><p>This notice serves to inform you that you have been placed on an academic improvement plan due to your falling below 70% required assignment completion. This improvement plan is not meant as a punitive measure but to help you on your journey to academic recovery.</p><p>As a Student on an <strong>Academic Improvement Plan</strong>, you are required to complete the following:</p><ol><li> Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a <strong> minimum</strong>  of 1 Code Review per week with Instruction Staff.</li><li>Attend <strong> all</strong>  scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <strong> {stack} stack</strong>  with Instructor {instructor} is the following:</p><ol><li>Status: <strong> Falling Behind </strong> </li><li>Assignment progress: <strong> {assignment_completion}%</strong> </li><li>Discussion progress: <strong> {discussion_completion}%</strong> </li></ol><p>To Return to <strong> Good Academic Standing</strong> , you must complete the following:</p><ol><li>Fulfill the requirements listed above.</li><li>Maintain a minimum of <strong> 70% </strong>required assignment completion.</li><li>Maintain a minimum of <strong>80% </strong>  discussion completion.</li></ol><p>Failure to return to Good Academic Standing can result in having to retake the stack, Academic Probation or dismissal from the program. If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact your {student_support}</p>";

    let old_templateOnPace = `<p>Hello {name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week. You may already be in contact with your instructor or student support regarding this matter.</p><p>This is an update of your progress during the <strong>{stack}</strong> stack with Instructor <strong>{instructor}</strong>. All students are to maintain above 80% in discussion completion and above 70% completion of mandatory assignments.</p><p>Current Standing: <strong>On Pace - Keep up the good work!</strong></p><ul><li>Assignments completion: {assignment_completion}%</li> <li>Discussion completion: {discussion_completion}%</li></ul></p><p> Failure to maintain the required percentage stated above throughout the rest of the stack may result in being placed on an Academic Improvement Plan.</p><p>Feel free to reach out to your {student_support}, for any questions or concerns.</p>`;

    let old_templateAcademicProbation = `<p>Hello {name},</p><p>This notice is to inform you that you have been placed on academic probation due to failing to meet the academic progress requirements as outlined in the catalog. Coding Dojo places students on academic probation with less than 50% required assignment submission. We would like to see you return to Good Academic Standing, and we support you as you take the steps to academic recovery.</p><p>As a Student on <strong>Academic Probation</strong>, you are required to complete the following:</p><ol><li>Contact your Student Support Manager and current instructor to schedule regular updates on their academic progress.</li><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a <strong>minimum</strong> of 1 Code Review per week with Instruction Staff.</li><li>Attend <strong>all</strong> scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <strong>{stack}</strong> stack with <strong>Instructor {instructor} </strong> is the following:</p><ul><li>Status: <strong>{status}</strong></li><li>Assignment progress: <strong>{assignment_completion}%</strong></li><li>Discussion progress: <strong>{discussion_completion}%</strong></li></ul><p>To Return to Academic Satisfactory Progress, you must complete the following:</p><ol><li>Fulfill the requirements listed above.</li><li>Maintain a minimum of <strong>50%</strong> required assignment completion.</li><li>Maintain a minimum of <strong>60%</strong> discussion completion.</li></ol><p>Failure to return to Good Academic Standing or Academic Satisfactory Progress can result in having to retake the stack, Academic Probation or dismissal from the program. Students who are placed on academic probation 3 times during the course of the program will then be placed under review and may be withdrawn from the program. </p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact your {student_support}</p>`;

    let new_templateAcademicProbation = `<p>Hello {name},</p><p>This notice is to inform you that you have been placed on Academic Probation due to failing to meet the academic progress requirements as outlined in the catalog and having made unsatisfactory progress. Coding Dojo places students on Academic Probation with less than 60% Core Assignment submission. We would like to see you return to Satisfactory Progress, and we support you as you take the steps to academic recovery.</p><p>As a Student on <b>Academic Probation</b>, you are required to complete the following:</p><ol><li>Contact your Student Support Manager and current instructor to schedule regular updates on their academic progress.</li><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a <b>minimum</b> of 1 Code Review per week with Instruction Staff.</li><li>Attend <b>all</b> scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <b>{stack}</b> stack with <b>Instructor {instructor}</b> is the following:</p><ul><li><b>Academic Probation</b></li><li>Assignment progress: <b>{assignment_completion}%</b></li><li>Discussion progress: {discussion_completion}%</li></ul><p>To return to Marginal Progress and no longer be on Academic Probation, you must complete the following:</p><ol><li>Fulfill the requirements listed above.</li><li>Maintain a minimum of <b>60%</b> Core assignment completion.</li><li>Maintain a minimum of <b>80%</b> discussion completion.</li></ol><p>Students who meet all Academic Probation requirements and/or fail to meet progression standards will be placed on an additional instance of Academic Probation at the following progression check during the stack. Students who fail to reach good standing by the end of the stack will be placed under Academic Review to determine the course of action, which may include: Retake of Stack, Withdrawal from Program, Program Transfer (Web Fundamentals Stack only). Three or more instances of unexcused academic probation on a student’s record will be grounds for academic dismissal from the program.</p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact our Student Support team at <a href="support@codingdojo.com">support@codingdojo.com</a>. A student’s status of Academic Probation may be lifted once the student returns to Marginal Progress standing (defined above) at minimum or is excused. A student is allowed one excused Academic Probation per stack</p>`

    let new_templateAssignmentWarning = `<p>Hello {name},</p><p>This notice serves to inform you that you have been placed on an <b>Academic Improvement Plan</b> due to your falling below 90% of Core assignment completion. This improvement plan is not meant as a punitive measure but to help you on your journey to academic recovery. </p><p>As a Student on an <b>Academic Improvement Plan</b>, it is recommended for you to complete the following:</p><ol><li>Schedule a 1:1 with the instructor to review progress and understanding of course material.</li><li>Attend a minimum of 1 Code Review per week with Instruction Staff.</li><li>Attend all scheduled appointments with Coding Dojo staff, including additional mandatory Code Reviews.</li></ol><p>For reference your current standing during your <b>{stack}</b> stack with <b>Instructor {instructor}</b> is the following:</p><ul><li><b>Marginal Progress - {status} </b></li><li>Assignment progress: <b>{assignment_completion}%</b></li><li>Discussion progress:  <b>{discussion_completion}%</b></li></ul><p>To Return to <b>Satisfactory Progress</b>, you must complete the following:</p><ol><li>Maintain a minimum of 90% Core assignment completion.</li><li>Maintain a minimum of 80% Discussion completion.</li></ol><p>Students who fail to meet all Academic Improvement Plan requirements and/or fail to meet progression standards will remain on an Academic Improvement Plan at the following progression check during the stack. Students who fail to reach satisfactory progress by the end of the stack will be placed under Academic Review to determine the course of action, which may include: Retake of Stack, Withdrawal from Program, Program Transfer (Web Fundamentals Stack only). A Student who falls below 60% of Core Assignment completion, as determined through regular review of student progress or meets other Academic Probation Criteria will be placed on Academic Probation. Three or more instances of unexcused academic probation on a student’s record will be grounds for academic dismissal from the program.</p><p>If you disagree with this decision and course of action, something has come up that affects your ability to progress with the program or should you have any questions regarding this letter or your academic standing, please contact our Student Support team at <a href="support@codingdojo.com">support@codingdojo.com</a>.</p>`

    let new_templateOnPace = `<p>Hello {name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week and is an update of your progress during the <b>{stack}</b> stack with <b>Instructor {instructor}</b>. All students are expected to maintain above 90% completion of Core Assignments and 80% completion of discussions</p><p>Current Standing:</p><ul><li><b>Satisfactory Progress - Keep up the good work!</b></li><li>Assignments completion: {assignment_completion}%</li><li>Discussion completion: {discussion_completion}%</li><p>*Failure to maintain at minimum 90% completion of Core Assignments at the following progression check will result in being placed on an Academic Improvement Plan. Falling below 60% completion of Core Assignments will result in being placed on Academic Probation. Students who fail to meet Satisfactory Progress by the end of the stack will be placed on Academic Review to determine the course of action, which may include: retake of stack, withdrawal from program or program transfer (Web Fundamentals stack only).</p><p>Feel free to reach out to your Student Support team at <a href="support@codingdojo.com">support@codingdojo.com</a>, for any questions or concerns.</p>`

    let new_templateDiscussionWarning = `<p>Hello {name},</p><p>This is a semi-automated message based on the data pulled from your student account at the beginning of this week. You may already be in contact with your instructor or student support regarding this matter.</p><p>This is a warning regarding your progress during the <b>{stack}</b> stack with <b>Instructor {instructor}</b>. All students are expected to maintain above 80% in Discussion Completion and above 90% completion of Core Assignments.</p><p>Current Standing: <b>Discussion Warning</b></p><ul><li>Discussion completion: <b>{discussion_completion}%</b></li><li>Assignments completion: <b>{assignment_completion}%</b></li></ul><p>A student who has not logged on to the Online Learning Platform and posted on their assigned Discussion for more than five (5) consecutive assigned sessions, is considered inactive. In this instance, the student will face termination from the bootcamp if they are unable to be reached by Coding Dojo staff. A student who has missed more than 20% of the required discussions by the end of the stack will be required to withdraw from the program. Campus staff may excuse up to 10% of a student’s missed discussions for special or mitigating circumstances outside the control of the student. In those cases, the circumstances must be provided, in writing, to campus staff as soon as possible.</p><p>Feel free to reach out to the Student Support Team at support@codingdojo.com if you have any questions.</p>`

    const leadInstructorEmail = "mtaylor@codingdojo.com";
    const studentSupportEmail = "proehl@codingdojo.com";
    const doNotSendList = ['POSTPONED(PPD)', 'DROPPED(D)', 'ROLLBACK(RB)', 'LIMBO(L)', 'PAUSED(P)', 'EXPELLED(E)', 'MOVEDFORWARD(MF)', 'SELFPACED(SP)', 'TEXAS(T)'];

    let subjectLine = '';
    let template = '';

    let statusCounts = {
        'ONPACE(OP)': 0,
        'FALLINGBEHIND(B)': 0,
        'DQWARNING(DTW)': 0,
        'APWARNING(APW)': 0,
        'PROBATION(AP)': 0,
    }

    if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com" || Session.getActiveUser().getEmail() == "program-status@codingdojo.com") {
        const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var classInfo = ss.getRange(ss.createTextFinder('INSTRUCTOR DETAILS').findNext().getRow() + 1, 2, 4, 1).getValues();
        const instructor = classInfo[0][0];
        const instructorEmail = classInfo[2][0];
        const stack = classInfo[3][0];

        //New Catalog vs Old Catalog 
        const catalog = ss.getRange(1, 3).getValue() == "NEW" ? "NEW" : "OLD"

        let i = 4
        let emailStudents = [];
        while (ss.getRange(i, 1).getBackground() !== "#000000") {
            if (ss.getRange(i, 1).getValue() != '') {

                let status = ss.getRange(i, 9).getValue();
                status = status.split(" ").join("");
                if (!doNotSendList.includes(status)) {
                    emailStudents.push(ss.getRange(i, 1, 1, 11).getValues());
                    statusCounts[status]++
                }
            }
            i++
        }

        let result = SpreadsheetApp.getUi().alert(`✉ Send Emails ${catalog}`, `You're about to send emails to ${emailStudents.length} students including \n ${prettyPrint(statusCounts)}`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

        if (result === SpreadsheetApp.getUi().Button.OK) {
            let emailCC = `${studentSupportEmail}, ${instructorEmail}`;
            for (let student of emailStudents) {
                let curr_student = student[0]
                let name = curr_student[0];
                let email = curr_student[2];
                let user_assignment = curr_student[9].toFixed(0);
                let discussion_completion = curr_student[10].toFixed(0);
                let status = curr_student[8];

                if (status === 'ON PACE (OP)') {
                    template = catalog == 'NEW' ? new_templateOnPace : old_templateOnPace;
                    subjectLine = `Coding Dojo - ${stack} - Weekly Progress Update`;
                } else if (status === 'DQ WARNING (DTW)') {
                    template = catalog == 'NEW' ? new_templateDiscussionWarning : old_templateDiscussionWarning;
                    subjectLine = `Coding Dojo - ${stack} - Discussions Topic Warning`;
                } else if (status === 'AP WARNING (APW)' || status === 'FALLING BEHIND (B)') {
                    template = catalog == 'NEW' ? new_templateAssignmentWarning : old_templateAssignmentWarning;
                    subjectLine = `Coding Dojo - ${stack} - Academic Improvement Plan`;
                    emailCC += `, ${leadInstructorEmail}`;
                } else if (status == 'PROBATION (AP)') {
                    template = catalog == 'NEW' ? new_templateAcademicProbation : old_templateAcademicProbation;
                    subjectLine = `Coding Dojo - ${stack} - Academic Probation`;
                    emailCC += `, ${leadInstructorEmail}`;
                    file = catalog == 'NEW' ? DriveApp.getFileById('1-G1f-HbbxyFGkXDBq9joJ9zAkIzDmJKH') : DriveApp.getFileById('1K-oge8Tef2XLvc42MyFqSJBovoBt_DjZ');
                }
                var message = template.replace("{name}", name).replace("{assignment_completion}", user_assignment).replace("{discussion_completion}", discussion_completion).replace("{stack}", stack).replace("{student_support}", studentSupportEmail).replace("{status}", status).replace("{instructor}", instructor);
                // TESTING
                if (Session.getActiveUser().getEmail() == "dadorno@codingdojo.com") {
                    emailCC = '';
                }
                if (status == 'PROBATION (AP)') {
                    MailApp.sendEmail(
                        email,
                        subjectLine,
                        message, {
                            htmlBody: message,
                            cc: emailCC,
                            attachments: [file.getAs(MimeType.PDF)]
                        }
                    );
                } else {
                    MailApp.sendEmail(email, subjectLine, message, {
                        htmlBody: message,
                        cc: emailCC
                    });
                }
            }
            ssPrint(`${emailStudents.length} emails were sent.`);
        } else {
            ssPrint("No emails were sent.");
        }
    } else {
        SpreadsheetApp.getUi().alert("You are not authorized to send emails.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

function prettyPrint(obj) {
    returnStr = ''
    for (let [key, value] of Object.entries(obj)) {
        returnStr += `${key}: ${value}\n`;
    }
    return returnStr
}

function ssPrint(str) {
    SpreadsheetApp.getActive().toast(str);
}