function assignEmployeesToNewProject(projectId) {
  const ss = SpreadsheetApp.openById('1m30iuAC4l4sbWMW9UBX_VTdNvZH5Vt_TvABaCUs6irw');
  const employeesSheet = ss.getSheetByName('Employees');
  const projectsSheet = ss.getSheetByName('Projects');

  if (!employeesSheet || !projectsSheet) {
    Logger.log('Employees or Projects sheet not found');
    return;
  }

  const employees = employeesSheet.getDataRange().getValues();
  const projects = projectsSheet.getDataRange().getValues();

  Logger.log('Project ID: ' + projectId);
  Logger.log('Projects data: ' + JSON.stringify(projects));

  const projectRow = projects.findIndex(row => row[0] === projectId);

  if (projectRow === -1) {
    Logger.log('Project not found');
    return;
  }

  const project = projects[projectRow];
  const projectName = project[1];
  const requiredSkills = project[5].split(',').map(skill => skill.trim());
  const numEmployeesRequired = project[6];
  const assignedEmployees = [];
  const employeeCalendarLinks = {};

  for (let j = 1; j < employees.length; j++) {
    const employee = employees[j];
    const employeeSkills = employee[5].split(',').map(skill => skill.trim());

    // Check if employee has required skills and is available
    if (employeeSkills.some(skill => requiredSkills.includes(skill)) && checkEmployeeAvailability(employee[7], project[3], project[4])) {
      assignedEmployees.push(employee[1]); // Add employee name to the assigned list
      employeeCalendarLinks[employee[1]] = employee[7]; // Store employee calendar link

      // Send notification email to assigned employee
      sendEmployeeAssignmentEmail(employee[1], employee[2], project[1], project[3], project[4]);

      if (assignedEmployees.length >= numEmployeesRequired) break;
    }
  }

  // Update the "Employees Assigned" column in the Projects sheet
  try {
    const assignedRange = projectsSheet.getRange(projectRow + 1, 8);
    Logger.log('Updating cell: ' + assignedRange.getA1Notation());
    assignedRange.setValue(assignedEmployees.join(','));
    SpreadsheetApp.flush(); // Ensure changes are committed
    Logger.log('Employees assigned updated in spreadsheet: ' + assignedEmployees.join(','));
  } catch (e) {
    Logger.log('Error updating spreadsheet: ' + e.message);
  }

  Logger.log('Employees assigned: ' + assignedEmployees.join(', '));

  // Add events to employee calendars
  for (const employee of assignedEmployees) {
    const calendarUrl = employeeCalendarLinks[employee];
    addProjectToEmployeeCalendar(employee, project, calendarUrl);

  }

  // Check for employee shortage and send email if needed
  if (assignedEmployees.length < numEmployeesRequired) {
    const insufficientEmployees = numEmployeesRequired - assignedEmployees.length;
    sendInsufficientAssignmentEmail('lawmeijun204@gmail.com', projectName, insufficientEmployees, assignedEmployees, numEmployeesRequired);
  }
}

function checkEmployeeAvailability(calendarUrl, startDate, endDate) {
  try {
    // Extract Calendar ID from the URL
    const calendarId = extractCalendarId(calendarUrl);
    Logger.log('Extracted Calendar ID: ' + calendarId);

    // Access the calendar using the Calendar ID
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      Logger.log('Calendar not found: ' + calendarId);
      return false;
    }

    // Check for events in the specified date range
    const events = calendar.getEvents(new Date(startDate), new Date(endDate));
    Logger.log('Events found: ' + events.length);
    return events.length === 0;
  } catch (e) {
    Logger.log('Error checking availability: ' + e.message);
    return false;
  }
}

function addProjectToEmployeeCalendar(employee, project, calendarUrl) {
  try {
    // Extract Calendar ID from the URL
    const calendarId = extractCalendarId(calendarUrl);
    Logger.log('Extracted Calendar ID for adding event: ' + calendarId);

    // Access the calendar using the Calendar ID
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      Logger.log('Calendar not found: ' + calendarId);
      return;
    }

    // Add the event to the calendar
    calendar.createEvent(
      project[1], // Event title: Project Name
      new Date(project[3]), // Start date
      new Date(project[4]), // End date
      { description: project[2] } // Description: Project Description
    );

    Logger.log('Event added to calendar for ' + employee);
  } catch (e) {
    Logger.log('Error adding event to calendar: ' + e.message);
  }
}

function extractCalendarId(calendarUrl) {
  const calendarIdMatch = calendarUrl.match(/src=([^&]+)/);
  if (!calendarIdMatch) {
    Logger.log('Calendar ID not found in URL: ' + calendarUrl);
    return 'lawmeijun204@gmail.com'; // Default Calendar ID
  }
  return decodeURIComponent(calendarIdMatch[1]);
}

// function sendInsufficientAssignmentEmail(ownerEmail, projectName, insufficientEmployees) {
//   const subject = `Insufficient Employee Assignment for Project: ${projectName}`;
//   var requiredEmployees = projectName[6];
//   var assignedEmployees = projectName[8] ? projectName[8].split(',').length : 0;

//   const body = `There are not enough employees assigned to the project "${projectName}".\n\nRequired Employees: ${projectName[6]}\nAssigned Employees: ${(projectName[8] ? projectName[8].split(',').length : 0)}\nInsufficient Employees: ${insufficientEmployees}\n\nPlease review the project and take necessary actions to fulfill the required number of employees.\n\nBest regards,\nThe Project Management System`;
//   GmailApp.sendEmail(ownerEmail, subject, body);
// }

function sendInsufficientAssignmentEmail(ownerEmail, projectName, insufficientEmployees, assignedEmployees, requiredEmployees) {
  const subject = `Insufficient Employee Assignment for Project: ${projectName}`;

  const template = HtmlService.createTemplateFromFile('insufficientEmail');

  template.projectName = projectName;
  template.requiredEmployees = requiredEmployees;
  template.assignedEmployees = assignedEmployees.length > 0 ? assignedEmployees.join(', ') : '-';;
  template.insufficientEmployees = insufficientEmployees;

  var message = template.evaluate().getContent();

  GmailApp.sendEmail(ownerEmail, subject, '', {htmlBody: message});
}

function sendEmployeeAssignmentEmail(employeeName, employeeEmail, projectName, startDate, endDate) {
  const subject = `Assigned to Project: ${projectName}`;

  // Load the HTML template as a text string
  const template = HtmlService.createTemplateFromFile('assignmentEmail');
  
  var formatStartDate = Utilities.formatDate(startDate, "GMT+8", "dd/MM/yyyy");
  var formatEndDate = Utilities.formatDate(endDate, "GMT+8", "dd/MM/yyyy");

  template.employeeName = employeeName;
  template.projectName = projectName;
  template.startDate = formatStartDate;
  template.endDate = formatEndDate;

  var message = template.evaluate().getContent();

  // Send the email with HTML content
  GmailApp.sendEmail(employeeEmail, subject, '', {htmlBody: message});
}


function getEmployeeEmail(employeeName) {
  const ss = SpreadsheetApp.openById('1m30iuAC4l4sbWMW9UBX_VTdNvZH5Vt_TvABaCUs6irw');
  const employeesSheet = ss.getSheetByName('Employees');
  const employees = employeesSheet.getDataRange().getValues();
  for (const row of employees) {
    if (row[1] === employeeName) {
      return row[7]; // column 7 contains email addresses
    }
  }
  return null;
}
