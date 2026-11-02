// Sample VTU / BMSIT Result Dataset Generator for 6th Semester ISE
function generateSampleResultData() {
  const subjects = [
    { code: '21IS61', title: 'Software Architecture & Design Patterns', credits: 4, faculty: 'Dr. S. N. Prasad' },
    { code: '21IS62', title: 'Full Stack Development', credits: 4, faculty: 'Prof. Kavitha R' },
    { code: '21CS63', title: 'Machine Learning & Applications', credits: 4, faculty: 'Dr. Mohan Kumar P' },
    { code: '21IS641', title: 'Cloud Computing (Prof. Elective I)', credits: 3, faculty: 'Prof. Divya Sharma' },
    { code: '21IS642', title: 'Advanced Cyber Security (Prof. Elective I)', credits: 3, faculty: 'Dr. Rajesh Gowda' },
    { code: '21IS651', title: 'Renewable Energy Sources (Open Elective)', credits: 3, faculty: 'Prof. Arunkumar H' },
    { code: '21ISL66', title: 'Machine Learning Laboratory', credits: 1.5, faculty: 'Dr. Mohan Kumar P & Prof. Sneha K' },
    { code: '21ISL67', title: 'Full Stack Development Laboratory', credits: 1.5, faculty: 'Prof. Kavitha R & Prof. Divya Sharma' }
  ];

  const firstNames = [
    'Aarav', 'Aditi', 'Akash', 'Ananya', 'Arjun', 'Bhavana', 'Chetan', 'Deepika', 'Darshan', 'Divya',
    'Gautam', 'Harshitha', 'Ishaan', 'Janani', 'Kiran', 'Kavya', 'Likhith', 'Manish', 'Meghana', 'Naveen',
    'Neha', 'Nikhil', 'Pooja', 'Pawan', 'Pragathi', 'Pranav', 'Priyanka', 'Rahul', 'Rakshitha', 'Rohan',
    'Sahana', 'Sameer', 'Shreya', 'Siddharth', 'Sinchana', 'Suhas', 'Suraj', 'Swathi', 'Tejas', 'Tanvi',
    'Tarun', 'Trisha', 'Varun', 'Vandana', 'Vikas', 'Vinay', 'Yashas', 'Yuktha', 'Chaitra', 'Karthik',
    'Megha', 'Praveen', 'Roopa', 'Sanjay', 'Shalini', 'Sunil', 'Usha', 'Vidyashree', 'Vijay', 'Yash'
  ];
  const lastNames = ['Kumar', 'Rao', 'Sharma', 'Patil', 'Gowda', 'Hegde', 'Murthy', 'Reddy', 'Shetty', 'Bhat'];

  const categories = ['CET', 'CET-SNQ', 'COMED-K', 'COMEDK', 'MGMT', 'CET-SC/ST', 'CET-CAT1', 'D-CET'];

  const rows = [];
  const studentCount = 60;

  for (let i = 1; i <= studentCount; i++) {
    const isLateral = i > 52;
    const usnNum = isLateral ? String(400 + (i - 52)).padStart(3, '0') : String(i).padStart(3, '0');
    const usn = isLateral ? `1BY23IS${usnNum}` : `1BY22IS${usnNum}`;
    const name = `${firstNames[(i - 1) % firstNames.length]} ${lastNames[(i * 3) % lastNames.length]}`;
    const category = categories[i % categories.length];

    // Ability level of student determines grades
    const ability = (i % 10 === 0) ? 'fail_risk' : (i % 7 === 0) ? 'topper' : (i % 3 === 0) ? 'good' : 'average';

    let totalMarksSum = 0;
    let earnedCreditsSum = 0;
    let registeredCreditsSum = 0;
    let hasFail = false;

    // Student takes 6 theory/open subjects and 2 labs
    const studentSubjects = [
      subjects[0],
      subjects[1],
      subjects[2],
      (i % 2 === 0) ? subjects[3] : subjects[4], // elective choice
      subjects[5],
      subjects[6],
      subjects[7]
    ];

    studentSubjects.forEach((sub, sIdx) => {
      let grade = 'A';
      let total = 75;

      if (ability === 'topper') {
        grade = sIdx % 2 === 0 ? 'O' : 'A+';
        total = Math.floor(88 + Math.random() * 10);
      } else if (ability === 'good') {
        grade = ['A+', 'A', 'B+'][sIdx % 3];
        total = Math.floor(72 + Math.random() * 15);
      } else if (ability === 'fail_risk') {
        if (sIdx === 2 || sIdx === 0) { // Failed in ML or S&D
          grade = 'F';
          total = Math.floor(25 + Math.random() * 14);
          hasFail = true;
        } else if (sIdx === 3 && i === 10) {
          grade = 'AB'; // Absent in 1 exam
          total = 0;
        } else if (sIdx === 1 && i === 20) {
          grade = 'I'; // Illness grade
          total = 38;
        } else if (sIdx === 4 && i === 30) {
          grade = 'DX'; // Detained
          total = 20;
          hasFail = true;
        } else {
          grade = ['C', 'P', 'B'][sIdx % 3];
          total = Math.floor(45 + Math.random() * 18);
        }
      } else {
        grade = ['A', 'B+', 'B', 'C', 'P'][sIdx % 5];
        total = Math.floor(52 + Math.random() * 26);
      }

      const gradePoints = { 'O': 10, 'A+': 9, 'A': 8, 'B+': 7, 'B': 6, 'C': 5, 'P': 4, 'F': 0, 'AB': 0, 'DX': 0, 'X': 0, 'I': 0, 'W': 0 }[grade] || 0;

      totalMarksSum += total;
      registeredCreditsSum += sub.credits;
      if (gradePoints > 0) earnedCreditsSum += (sub.credits * gradePoints);

      rows.push({
        usn,
        name,
        code: sub.code,
        title: sub.title,
        grade,
        sem: 6,
        total,
        credits: sub.credits,
        category,
        sgpa: 0, // calculated below
        cgpa: 0
      });
    });

    const calculatedSgpa = registeredCreditsSum ? +(earnedCreditsSum / registeredCreditsSum).toFixed(2) : 0;
    const calculatedCgpa = hasFail ? +(calculatedSgpa * 0.9).toFixed(2) : +(Math.min(9.9, calculatedSgpa + 0.15)).toFixed(2);

    // Apply calculated SGPA/CGPA to student's rows
    rows.forEach(r => {
      if (r.usn === usn) {
        r.sgpa = calculatedSgpa;
        r.cgpa = calculatedCgpa;
      }
    });
  }

  return {
    rows,
    facultyMap: {
      '21IS61': 'Dr. S. N. Prasad',
      '21IS62': 'Prof. Kavitha R',
      '21CS63': 'Dr. Mohan Kumar P',
      '21IS641': 'Prof. Divya Sharma',
      '21IS642': 'Dr. Rajesh Gowda',
      '21IS651': 'Prof. Arunkumar H',
      '21ISL66': 'Dr. Mohan Kumar P & Prof. Sneha K',
      '21ISL67': 'Prof. Kavitha R & Prof. Divya Sharma'
    }
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { generateSampleResultData };
}
