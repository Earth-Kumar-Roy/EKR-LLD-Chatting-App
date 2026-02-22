# EKR LLD Chatting App 💬

A robust Low-Level Design (LLD) messaging application featuring real-time communication and user presence tracking.

---

## 📝 Description

**EKR LLD Chatting App** is designed to demonstrate core messaging system principles including:
* **Real-time Messaging:** Instant communication between users.
* **Online/Offline Status:** Live indicators showing user availability.
* **Read Receipts:** Visual confirmation for "Seen" and "Delivered" statuses.
* **Text-Based Interface:** A clean, distraction-free environment for conversation.

---

## 💬 Features

* **Full Real-time Messaging:** Messages appear instantly without page reloads.
* **Status Indicators:** Know exactly when your friends are online or offline.
* **Seen/Delivered Receipts:** Track the lifecycle of every message sent.
* **User Discovery:** Simply press the **+** button to start a new chat with any user.

---

## 🚀 Getting Started

1.  **Access the App:** Open the [Deployment Link](https://script.google.com/macros/s/AKfycbw7Ax6WycXZF9f2Hpbo6xtrrEuM65vvZ-vJq5HC917Qv3AjitfaXNaKr8g-vAUz3QA/exec).
2.  **Authentication:** Log in using the provided test credentials.
3.  **Find Users:** Tap the **+** icon to search for and add users to your list.
4.  **Connect:** Select a user and start your conversation instantly.

---

## 🛠️ How It Works

The application utilizes a **Google Apps Script** backend to manage data flow between users. 
* **Frontend:** Lightweight HTML/JS for the UI.
* **Backend:** `doGet()` and `google.script.run` functions handle server-side logic.
* **Data Persistence:** User credentials and message history are securely managed via Google Sheets.

---

**Deployment Link:** [Click here to view the live project](https://script.google.com/macros/s/AKfycbw7Ax6WycXZF9f2Hpbo6xtrrEuM65vvZ-vJq5HC917Qv3AjitfaXNaKr8g-vAUz3QA/exec)
