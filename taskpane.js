Office.onReady(() => {
    document.getElementById("scoreButton").onclick = getProjectRelevance;
});

async function getProjectRelevance() {
    const subject = Office.context.mailbox.item.subject;
    const bodyPromise = new Promise((resolve) => {
        Office.context.mailbox.item.body.getAsync("text", (result) => {
            resolve(result.value);
        });
    });

    const body = await bodyPromise;

    // Call your backend API
    const response = await fetch("https://your-api-url.com/score-email", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            email_subject: subject,
            email_body: body,
            list_name: "Project List Name"
        })
    });

    const projectScores = await response.json();

    // Sort scores and display the most relevant project
    projectScores.sort((a, b) => b.score - a.score);
    const mostRelevantProject = projectScores[0];

    document.getElementById("output").innerHTML = `
        Most Relevant Project: ${mostRelevantProject.project_title} (Score: ${mostRelevantProject.score})
    `;
}
