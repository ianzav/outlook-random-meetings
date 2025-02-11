// Outlook Add-in for Scheduling Random Recurring Meetings

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("scheduleButton").onclick = scheduleMeetings;
        document.getElementById("addMeetingButton").onclick = addMeetingEntry;
    }
});

function addMeetingEntry() {
    const container = document.getElementById("meetingsContainer");
    const entry = document.createElement("div");
    entry.classList.add("meeting-entry");
    entry.innerHTML = `
        <input type="date" class="meeting-date" required>
        <input type="time" class="meeting-time" required>
        <button type="button" class="remove-meeting">Remove</button>
    `;
    entry.querySelector(".remove-meeting").onclick = () => entry.remove();
    container.appendChild(entry);
}

async function scheduleMeetings() {
    try {
        const meetingDetails = getMeetingDetailsFromUI();
        for (const meeting of meetingDetails) {
            await createMeeting(meeting);
        }
        showNotification("Meetings Scheduled Successfully!");
    } catch (error) {
        showNotification("Error: " + error.message);
    }
}

function getMeetingDetailsFromUI() {
    const meetings = [];
    document.querySelectorAll(".meeting-entry").forEach(entry => {
        const date = entry.querySelector(".meeting-date").value;
        const time = entry.querySelector(".meeting-time").value;
        meetings.push({ start: new Date(`${date}T${time}`), subject: "Custom Meeting" });
    });
    return meetings;
}

async function createMeeting(meeting) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.subject.setAsync(meeting.subject, asyncResult => {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Failed to set subject"));
            } else {
                Office.context.mailbox.item.start.setAsync(meeting.start, asyncResult2 => {
                    if (asyncResult2.status !== Office.AsyncResultStatus.Succeeded) {
                        reject(new Error("Failed to set start time"));
                    } else {
                        resolve();
                    }
                });
            }
        });
    });
}

function showNotification(message) {
    const notificationArea = document.getElementById("notification");
    notificationArea.innerText = message;
}
