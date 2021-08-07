const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
function getStudentData() {
  const catalog= ss.getRange(1,3).getValue() == "NEW" ? "NEW" : "OLD"
  let student_array=[];
  let week=ss.getRange(2,10).getValue().toString()
  let assignment_column=ss.getRange('L2:ZZ2').createTextFinder(week).findNext().getColumn()
  let discussion_column=ss.getRange('L2:ZZ2').createTextFinder(week).findNext().getColumn()+1

  let missing_student_list=''

  let page=0;
  let filter_data=getFilterData()
  let response = apiCall(page, filter_data);

  while(response[1] === '/'){
    response= response.split('\n')[29];
    students=response.substring(response.indexOf('[')+1,response.lastIndexOf(']')).split('}},{');
    let count=0;
    let temp_student='';

    for(let student of students){
      if (count === 0){
        temp_student=student+'}}';
      }
      else if (count === students.length-1){
        temp_student='{'+ student;
      }
      else{
        temp_student='{'+ student + '}}';
      }
      student_array.push(JSON.parse(temp_student));
      count++;
    }

    page++
    response=apiCall(page, filter_data)
  }
  
  for(let student of student_array){
    let email=student.student[2].email
    Logger.log(email)
    if(catalog == 'NEW'){
      var required_assignments=student.mandatory.substring(0,student.mandatory.indexOf('/'))
    }
    else{
      var required_assignments=parseInt(student.mandatory.substring(0,student.mandatory.indexOf('/')))+parseInt(student.practice.substring(0,student.practice.indexOf('/')))
    }
    let discussions=student.total_discussion.substring(0,student.total_discussion.indexOf('/'))
    let email_search=ss.getRange('C:C').createTextFinder(email).findNext()
    if(email_search){
      let row=email_search.getRow()
      ss.getRange(row,assignment_column).setValue(required_assignments)
      ss.getRange(row,discussion_column).setValue(discussions)
    }
    else{
      missing_student_list+=`${email} \n`
    }
  }

  ssPrint(missing_student_list)
}

function apiCall(page, filter_data){
    let options={
    "headers": {
      'Accept': '*/*',
      'X-CSRF-Token': 'JvMplIwjnsWwnkBHGvfCaihksg/uX93k5uTRfYOz0J4=',
      'Cookie': 'intercom-id-rpf3h5c2=59b0039c-2ab6-46e3-8849-d18d754a7a0b; intercom-session-rpf3h5c2=; _boomyeah_2_session=S2pzTGJqK1VXL1B4QXM2RkNpY01KUFRxaHo1Q29YQnlrWmFpUWJvL2RHK3Ryb3kxa0FzYW9JRDVUYzJxaHFMZmFsR3Z2SHNYbVJKN3hmUGUzcEREUlhhODFLTFdpRks3QXMzeW9CeDV3aCtPSTJhSWdBM2JrVllqZ3p6QmdYNjBpOENxSUdCbG11cE42azdVL21CK1R0TTRaaVRYeFhKVzdubzI0RkpLVEwwUTl0RTZhcXRHUjdtV1Q5a0ZnQXREYVpHeWxaYWhidjg3N0tQcmFZdW1zaG5GbmVFdnZ0bHE2dVE3dVArUldGNXUxbnF0NmN5N0s5QjlhMThtK3dMZndkYW5sQVQyMUZlZkZBTE4zdjBFZEovWEJYWUQ5TmJRdVNGZVhucXZSb2krYWd0U0xkM0U0ZERPT3ROb1JvZlVBQ2N4YVB2alEwaHVwU3ZGbnozQlBFalduYlJsTEhKeVU2ZkhhZUFFMy9tWTZVOGl2aEM5RDVyZTh3cUNhcTVFbVZRdkNjcVRkMFQ3bHdyb1N4ZHkzT3NaR0doUEpYZGdYdWpHL3BEREFZYW4yY0F2dUFVU2t6ZmZhbDNCU3NpZjkwTmovNmRucUNqQTB5NVlyT2FSaVRPZzBSejh0dGQ3RUdlQWNHVlMzY3hicjA0QUJ5bFBOdnpoN1J5b0M5YlBnOTB0M0QwYy9sYmVtZXhDRitFVlo2TXNWS20vSExNaW8xTGpzSVFPdjJmdk0xT3RjeWpXcHpwaVo0bi9Eekc3cW92RXFOaHlyQWEvVURCZEQzSE5HUFp3UWVQSDc5ci9vc2VCV1gvSkErRG95cFVZYTNoSndiRlptcnpzV3RTalVUU2xwTitBeWt0VXk1cFdFWVp6T1E9PS0tRjdIOCtOL2NYeUVPN0RJYXhHVi96UT09--ca9d7d7a35d29d8e90564e85824dd7100bebe1e4'
    },
    "method": 'post',
    "payload": {
        'program_type_id': '27',
        'location_id': '0',
        'cohort_status': '2',
        'searched_student': '',
        'stack_id': filter_data['stack_id'],
        'stack_year': filter_data['stack_year'],
        'stack_start_date': filter_data['stack_start_date'],
        'location_id': '0',
        'instructor_id': '0',
        'page':page
    },
  };

  return UrlFetchApp.fetch('http://learn.codingdojo.com/filter_new_admin_progress',options).getContentText();
}

function getFilterData(){
  let classInfo=ss.getRange(ss.createTextFinder('INSTRUCTOR DETAILS').findNext().getRow() + 1, 2, 5, 1).getValues()
  let instructor_name = `${classInfo[0][0]} ${classInfo[1][0]}`.toLowerCase()
  let stack = classInfo[3][0].toLowerCase()
  let stack_start_date=Utilities.formatDate(classInfo[4][0], "PST", "yyyy-MM-dd")
  let stack_year=`${classInfo[4][0].getFullYear()}`

  let instructor_ids={
    'drew adorno': '264',
    'michael  clinkscales': '265',
    'chris juarez': '398',
  }

  let stack_ids={
    'webfun': '99',
    'python': '4',
    'p&a': '9'
  }

  return {
    'instructor_id':instructor_ids[instructor_name],
    'stack_id': stack_ids[stack],
    'stack_start_date': stack_start_date,
    'stack_year':stack_year
  }
}

function ssPrint(str){
  SpreadsheetApp.getActive().toast(str);
}
