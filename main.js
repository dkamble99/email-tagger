async function getEmailDetails() {
    Office.context.mailbox.item.subject.getAsync((subjectResult) => {
        Office.context.mailbox.item.body.getAsync("text", async (bodyResult) => {
            const emailSubject = subjectResult.value;
            const emailBody = bodyResult.value;

            // Call the backend API
            const response = await fetch("https://aa01-46-18-28-243.ngrok-free.app/score-email/", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    email_subject: emailSubject,
                    email_body: emailBody,
                    list_name: "Lista dei Progetti"  // Adjust the list name if necessary
                })
            });

            const projectScores = await response.json();
            displayScores(projectScores);
        });
    });
}

function displayScores(scores) {
    const output = document.getElementById("score-output");
    output.innerHTML = scores.map(score => 
        `<div><strong>Project:</strong> ${score.project_title}<br><strong>Score:</strong> ${score.score}</div><hr>`
    ).join("");
}
