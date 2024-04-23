document.querySelector('#certificateForm').addEventListener('submit', handleFormSubmit);

function handleFormSubmit(e)
{
	e.preventDefault()
	let student_id = document.getElementById("studentId").value; // replace with your student id
	let url = 'http://127.0.0.1:8000/send_pdf'; // replace with your FastAPI server url

	let data = { student_id: student_id };

	fetch(url, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json'
		},
		body: JSON.stringify(data)
	})
		.then(response => response.blob())
		.then(blob =>
		{
			var url = window.URL.createObjectURL(blob);
			var a = document.createElement('a');
			a.href = url;
			a.download = "Certificate - " + student_id + '.pdf';
			document.body.appendChild(a);
			a.click();
			a.remove();
		})
		.catch((error) =>
		{
			console.error('Error:', error);
		});
}