function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.trim());
}

function validateMigrationData(data) {
  const errors = [];

  if (!data.jlid || data.jlid.trim() === '') {
    errors.push('JLID is required');
  }

  if (!data.learner || data.learner.trim() === '') {
    errors.push('Learner name is required');
  }

  if (!data.newTeacher || data.newTeacher.trim() === '') {
    errors.push('New teacher is required');
  }

  if (!data.course || data.course.trim() === '') {
    errors.push('Course is required');
  }

  if (!data.classSessions || data.classSessions.length === 0) {
    errors.push('At least one class session (Day and Time) is required');
  } else {
    data.classSessions.forEach((session, index) => {
      if (!session.day || session.day.trim() === '') {
        errors.push(`Class Day for session ${index + 1} is required`);
      }
      if (!session.time || !session.time.match(/^\d{1,2}:\d{2}\s(AM|PM)$/i)) {
        errors.push(`Class Time for session ${index + 1} is invalid or missing. Expected format HH:MM AM/PM`);
      }
    });
  }

  if (!data.clsManager || data.clsManager.trim() === '') {
    errors.push('CLS Manager is required');
  }

  if (!data.jetGuide || data.jetGuide.trim() === '') {
    errors.push('JetGuide is required');
  }

  if (!data.startDate || data.startDate.trim() === '') {
    errors.push('Start Date is required');
  }

  if (!data.migrationType || !['Mid-Course', 'New Assignment'].includes(data.migrationType)) {
    errors.push('Valid migration type is required');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

function validateRequiredFields(data, requiredFields) {
  if (!data) return "No data provided.";
  
  const missing = [];
  requiredFields.forEach(field => {
    if (!data[field] || String(data[field]).trim() === '') {
      missing.push(field);
    }
  });

  if (missing.length > 0) {
    return `Missing required fields: ${missing.join(', ')}`;
  }
  return null; // No errors
}
