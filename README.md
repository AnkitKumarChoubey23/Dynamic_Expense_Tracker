# Dynamic_Expense_Tracker
A dynamic, emoji-enhanced Python expense tracker that takes real-time user input, analyzes category-wise spending, compares against previous month benchmarks, and exports a recruiter-ready Excel report with styled analytics and total expense summary — all from the command line.

🚀 Features
• 	Dynamic Category Handling: Add any number of expense categories on the fly.
• 	Auto-Adjusted Columns: Excel output adapts to category count and label length.
• 	Analytics Built-In:
• 	Total, Min, Max per category
• 	Overall expense summary
• 	Previous month comparison
• 	Excel Export:
• 	Styled headers and totals
• 	Clear formatting for recruiter readability
• 	Saves to correct Desktop path across environments
• 	Robust Error Handling:
• 	Input validation
• 	File overwrite protection
• 	OS-aware path detection

📦 Requirements
pip install pandas openpyxl

🛠️ How to Run
Follow the prompts to enter your expenses. The script will:
• 	Save your data to an Excel file on your Desktop
• 	Auto-style the sheet for clarity
• 	Compare with previous month if available

📊 Sample Output
| Category | Min | Max | Total | Δ vs Last Month  | 
| Food | ₹50 | ₹300 | ₹1200    |      +₹200       | 
| Travel | ₹100 | ₹500 | ₹1800 |      -₹400       | 
| Rent | ₹5000 | ₹5000 | ₹5000 |        0         | 

📁 File Output
- Filename: Expense_Report_August_2025.xlsx
- Location: Automatically saved to your Desktop
- Format: Styled Excel with analytics and comparison

🧠 Behind the Scenes
Built with:
- pandas for data handling
- openpyxl for Excel styling
- Modular functions for input, analytics, and export
- OS-aware path detection for seamless saving

📌 Why This Project?
Designed to showcase:
- Real-world Python utility
- Recruiter-friendly output
- Iterative polish and robust error handling
- Visual clarity and user-centric design


🧑‍💻 Author
Ankit Kumar
B.E. Computer Science | SVCE Bangalore
Passionate about building polished, real-world tools with Python, UI/UX sensibility, and recruiter appeal.





