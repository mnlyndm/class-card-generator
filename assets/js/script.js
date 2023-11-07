document.getElementById('generate-button').addEventListener('click', function() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];

  if (file) {
    const reader = new FileReader();

    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const studentData = XLSX.utils.sheet_to_json(sheet);

      const classCards = document.getElementById('class-cards');
      classCards.innerHTML = '';

      studentData.forEach(student => {
        const lastName = student['Last Name'] || '';
        const firstName = student['First Name'] || '';
        const middleName = student['Middle Name'] || '';
        const level = student.Level || '';
        const course = student.Course || '';
        const subject = student.Subject || '';
        const professor = student.Professor || '';
        const day = student.Day || '';
        const time = student.Time || '';
        const rating = student['Final Rating'] || '';

        const card = `
          <div class="bg-white p-4 rounded shadow-lg mb-4">
            <h2 class="text-xl font-semibold">${lastName}, ${firstName} ${middleName}</h2>
            <p>Level: ${level}</p>
            <p>Course: ${course}</p>
            <p>Subject: ${subject}</p>
            <p>Professor: ${professor}</p>
            <p>Day: ${day}</p>
            <p>Time: ${time}</p>
            <p>Final Rating: ${rating}</p>
          </div>
        `;
        classCards.innerHTML += card;
      });
    };

    reader.readAsArrayBuffer(file);
  } else {
    alert('Please select an Excel file.');
  }
});
