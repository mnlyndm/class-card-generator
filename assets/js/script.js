document.getElementById('generate-button').addEventListener('click', function () {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];

  if (file) {
    const reader = new FileReader();

    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const studentData = XLSX.utils.sheet_to_json(sheet);

      const classCards = document.getElementById('class-cards');
      classCards.innerHTML = '';

      studentData.forEach(student => {
        const lastName = student['Last Name'] || '&nbsp;';
        const firstName = student['First Name'] || '&nbsp;';
        const middleName = student['Middle Name'] || '&nbsp;';
        const level = student.Level || '&nbsp;';
        const course = student.Course || '&nbsp;';
        const subject = student.Subject || '&nbsp;';
        const professor = student.Professor || '&nbsp;';
        const day = student.Day || '&nbsp;';
        const time = student.Time || '&nbsp;';
        const rating = student['Final Rating'] || '&nbsp;';

        const card = `
        <div class="flex flex-col justify-start w-full items-center bg-white border border-gray-400 p-1 mt-2 rounded-sm mb-4" style="width: 336px; height: 192px; ">
        <div>
          <img src="./assets/svg/ollclogo.svg" class="w-40" alt="">
        </div>
        <div style="color: #0D1158; font-size: 8px;">
          __ Semester, 20__-20__
        </div>
        <div class=" font-bold" style="color: #0D1158; font-size: 12px;">
          CLASS CARD
        </div>
        <div class=" w-full px-2 text-xs mt-1 inline-flex justify-start items-center gap-4" style="color: #0D1158; font-size: 8px;">
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${lastName} 
            </div>
            <div>
              Surname
            </div>
          </div>
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${firstName} 
            </div>
            <div>
              First Name
            </div>
          </div>
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${middleName} 
            </div>
            <div>
              Middle Name
            </div>
          </div>
        </div>
        <div class=" w-full px-2 text-xs mt-1 inline-flex justify-start items-center gap-4" style="color: #0D1158; font-size: 8px;">
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${level} 
            </div>
            <div>
              Level
            </div>
          </div>
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${course} 
            </div>
            <div>
              Course
            </div>
          </div>
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${subject} 
            </div>
            <div>
              Subject
            </div>
          </div>
        </div>
        <div class=" w-full px-2 text-xs mt-1 inline-flex justify-start items-center gap-4" style="color: #0D1158; font-size: 8px;">
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${professor} 
            </div>
            <div>
              Professor
            </div>
          </div>
          <div class="text-center w-full">
            <div class="flex justify-center items-center gap-1">
              <div>
                <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
                  ${day}
                </div>
                <div>
                  Day
                </div>
              </div>
              <div>
                <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
                  ${time} 
                </div>
                <div>
                  Time
                </div>
              </div>
            </div>
          </div>
          <div class="text-center w-full">
            <div class="border border-solid text-black" style="border-color: #0D1158; padding-left: 4px; padding-right: 4px;">
              ${rating} 
            </div>
            <div>
              Final Rating
            </div>
          </div>
        </div>
      <div>
        `;

        const cardContainer = document.createElement('div');
        cardContainer.className = 'card-container';

        cardContainer.innerHTML = card;

        const downloadButton = document.createElement('button');
        downloadButton.innerHTML = 'Download Card';
        downloadButton.addEventListener('click', function () {
          downloadCardAsImage(cardContainer);
        });

        cardContainer.appendChild(downloadButton);

        classCards.appendChild(cardContainer);
      });
    };

    reader.readAsArrayBuffer(file);
  } else {
    alert('Please select an Excel file.');
  }
});

function downloadCardAsImage(cardContainer) {
  html2canvas(cardContainer, {
    onclone: function (clone) {
      clone.querySelector('.card-container button').style.display = 'none';
    }
  }).then(canvas => {
    const imageData = canvas.toDataURL('image/png');

    const downloadLink = document.createElement('a');
    downloadLink.href = imageData;
    downloadLink.download = 'card.png';
    downloadLink.click();
  });
}
