// js for responsive nav
let menuicn = document.querySelector(".menu-symbol");
let nav = document.querySelector(".navcontainer");

menuicn.addEventListener("click", () =>
{
	nav.classList.toggle("navclose");
});

// js for handling formdata
document.querySelector('#imageUploadForm').addEventListener('submit', handleFormSubmit);

function handleFormSubmit(e)
{
	e.preventDefault(); // Prevent the default form submission behavior

	const formData = new FormData(e.target);
	const file = document.getElementById('imageUploadInput').files[0];
	formData.append('file', file);

	let hemlo = fetch('http://127.0.0.1:8000/upload', {
		method: 'POST',
		headers: {
			'accept': 'application/json'
		},
		body: formData
	})
		.then(response =>
		{
			if (!response.ok)
			{
				throw new Error('Network response was not ok');
			}
			return response.json();
		})
		.then(data =>
		{
			console.log(data);
			console.log("🚀 ~ marks:", data);

			// Create and append the result container
			const newDiv = document.createElement("div");
			newDiv.classList.add("result-container");
			const heading = document.createElement("h3");
			heading.textContent = "Following Information where recorded";
			const marksDiv = document.createElement("div");
			marksDiv.textContent = `Marks Scored: ${data.marks}`;
			const regNumDiv = document.createElement("div");
			regNumDiv.textContent = `Registration Number: ${data.student_id}`;
			const subjectDiv = document.createElement("div");
			subjectDiv.textContent = `Subject Code: ${data.subject_code}`;
			newDiv.appendChild(heading);
			newDiv.appendChild(marksDiv);
			newDiv.appendChild(regNumDiv);
			newDiv.appendChild(subjectDiv);
			document.querySelector("#imageUploadForm").appendChild(newDiv);
		})
		.catch(error =>
		{
			console.error('There was a problem with your fetch operation:', error);
		});
}