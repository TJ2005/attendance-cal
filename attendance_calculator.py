import pdfplumber
import pandas as pd
import re
from datetime import datetime
from pathlib import Path
import sys

class AttendanceCalculator:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.attendance_data = []
        self.subjects = {}
        self.student_name = ""
        self.sap_id = ""
        self.program = ""
        self.batch = ""
        self.report_date = datetime.now().strftime("%B %d, %Y")
        
    def extract_data(self):
        """Extract attendance data from PDF"""
        print(f"Processing: {Path(self.pdf_path).name}")
        print("=" * 80)
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                # Extract student info from first page
                if len(pdf.pages) > 0:
                    first_page_text = pdf.pages[0].extract_text()
                    lines = first_page_text.split('\n')
                    
                    for i, line in enumerate(lines):
                        if 'Name' in line and ':' in line:
                            self.student_name = line.split(':', 1)[1].strip()
                        elif 'SAP ID' in line or 'Student ID' in line:
                            if ':' in line:
                                self.sap_id = line.split(':', 1)[1].strip()
                            elif i + 1 < len(lines):
                                self.sap_id = lines[i + 1].strip()
                        elif 'Program' in line or 'Programme' in line:
                            if ':' in line:
                                self.program = line.split(':', 1)[1].strip()
                        elif 'Batch' in line:
                            if ':' in line:
                                self.batch = line.split(':', 1)[1].strip()
                
                for page_num, page in enumerate(pdf.pages, 1):
                    tables = page.extract_tables()
                    
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        
                        # Skip header row
                        for row in table[1:]:
                            if len(row) >= 6 and row[0] and row[0].strip().isdigit():
                                try:
                                    self.attendance_data.append({
                                        'sr_no': int(row[0].strip()),
                                        'course': row[1].strip() if row[1] else '',
                                        'date': row[2].strip() if row[2] else '',
                                        'start_time': row[3].strip() if row[3] else '',
                                        'end_time': row[4].strip() if row[4] else '',
                                        'attendance': row[5].strip() if row[5] else ''
                                    })
                                except (ValueError, IndexError) as e:
                                    continue
                
                print(f"Extracted {len(self.attendance_data)} lecture records")
                if self.student_name:
                    print(f"Student: {self.student_name}")
                if self.sap_id:
                    print(f"SAP ID: {self.sap_id}\n")
                return True
                
        except Exception as e:
            print(f"Error processing PDF: {e}")
            return False
    
    def clean_course_name(self, course_name):
        """Extract base course name by removing course type and section info"""
        # Remove common patterns like T1, P1, batch info, etc.
        clean_name = re.sub(r'[TP]\d+\s*-?\s*', '', course_name)
        clean_name = re.sub(r'(BT|BTech|B\.Tech).*', '', clean_name)
        clean_name = re.sub(r'(Cyber|OE\d+|BTMT\d+|MBA).*', '', clean_name)
        clean_name = clean_name.strip()
        
        # Handle specific course names
        if 'AI and ML' in clean_name or 'Cybersecurity' in course_name and 'AI' in course_name:
            return 'AI and ML for Cybersecurity'
        elif 'Network Security' in clean_name:
            return 'Network Security'
        elif 'Visual Analytics' in clean_name:
            return 'Visual Analytics'
        elif 'Software Engineering' in clean_name:
            return 'Software Engineering'
        elif 'Cybersecurity Fundamentals' in clean_name:
            return 'Cybersecurity Fundamentals'
        elif 'Forensic' in clean_name:
            return 'Introduction to Forensic Science'
        elif 'Drone' in clean_name:
            return 'Drone Technology'
        
        return clean_name
    
    def calculate_subject_attendance(self):
        """Calculate attendance percentage for each subject"""
        if not self.attendance_data:
            print("No data to calculate. Please run extract_data() first.")
            return
        
        # Group by subject
        for record in self.attendance_data:
            subject = self.clean_course_name(record['course'])
            
            if subject not in self.subjects:
                self.subjects[subject] = {
                    'total': 0,
                    'present': 0,
                    'absent': 0,
                    'absent_dates': [],
                    'lectures': []
                }
            
            self.subjects[subject]['total'] += 1
            self.subjects[subject]['lectures'].append(record)
            
            if record['attendance'] == 'P':
                self.subjects[subject]['present'] += 1
            elif record['attendance'] == 'A':
                self.subjects[subject]['absent'] += 1
                self.subjects[subject]['absent_dates'].append(record['date'])
    
    def generate_html_report(self, output_file='attendance_report.html'):
        """Generate a minimal black and white HTML attendance report"""
        if not self.subjects:
            print("No data available to generate report.")
            return
        
        # Calculate overall statistics
        total_lectures = sum(s['total'] for s in self.subjects.values())
        total_present = sum(s['present'] for s in self.subjects.values())
        total_absent = sum(s['absent'] for s in self.subjects.values())
        overall_percentage = (total_present / total_lectures * 100) if total_lectures > 0 else 0
        
        # Sort subjects alphabetically
        sorted_subjects = sorted(self.subjects.items(), key=lambda x: x[0])
        
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Attendance Report</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            background: #ffffff;
            color: #000000;
            line-height: 1.6;
            padding: 40px 20px;
        }}
        
        .container {{
            max-width: 900px;
            margin: 0 auto;
            animation: fadeIn 0.6s ease-in;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 60px;
            padding-bottom: 30px;
            border-bottom: 1px solid #e0e0e0;
            animation: slideDown 0.8s ease-out;
        }}
        
        @keyframes slideDown {{
            from {{ opacity: 0; transform: translateY(-30px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .header h1 {{
            font-size: 2.2em;
            font-weight: 300;
            letter-spacing: -0.5px;
            margin-bottom: 20px;
        }}
        
        .student-info {{
            background: #fafafa;
            padding: 25px;
            margin-bottom: 50px;
            border-left: 3px solid #000;
            animation: fadeIn 0.8s ease-in 0.2s both;
        }}
        
        .student-info h2 {{
            font-size: 1.1em;
            font-weight: 600;
            margin-bottom: 15px;
            letter-spacing: 0.5px;
        }}
        
        .info-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }}
        
        .info-item {{
            display: flex;
            flex-direction: column;
        }}
        
        .info-label {{
            font-size: 0.75em;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #666;
            margin-bottom: 5px;
        }}
        
        .info-value {{
            font-size: 1em;
            font-weight: 500;
        }}
        
        .overall-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
            gap: 20px;
            margin-bottom: 60px;
            animation: fadeIn 0.8s ease-in 0.4s both;
        }}
        
        .stat-box {{
            text-align: center;
            padding: 20px;
            border: 1px solid #e0e0e0;
            transition: all 0.3s ease;
        }}
        
        .stat-box:hover {{
            transform: translateY(-5px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }}
        
        .stat-label {{
            font-size: 0.7em;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: #666;
            margin-bottom: 10px;
        }}
        
        .stat-value {{
            font-size: 2.5em;
            font-weight: 200;
            line-height: 1;
        }}
        
        .subjects-section {{
            margin-top: 40px;
        }}
        
        .section-title {{
            font-size: 1.5em;
            font-weight: 300;
            margin-bottom: 30px;
            letter-spacing: -0.5px;
        }}
        
        .subject-card {{
            margin-bottom: 40px;
            border: 1px solid #e0e0e0;
            padding: 30px;
            transition: all 0.3s ease;
            animation: fadeIn 0.6s ease-in both;
            position: relative;
        }}
        
        .subject-card:hover {{
            box-shadow: 0 5px 20px rgba(0,0,0,0.08);
        }}
        
        .subject-header {{
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid #f0f0f0;
        }}
        
        .subject-name {{
            font-size: 1.3em;
            font-weight: 400;
        }}
        
        .subject-percentage {{
            font-size: 2em;
            font-weight: 200;
        }}
        
        .subject-stats {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 25px;
        }}
        
        .stat-item {{
            text-align: center;
            padding: 15px;
            background: #fafafa;
        }}
        
        .stat-item-label {{
            font-size: 0.7em;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #999;
            margin-bottom: 8px;
        }}
        
        .stat-item-value {{
            font-size: 1.8em;
            font-weight: 300;
        }}
        
        .calculator-section {{
            margin-top: 25px;
            padding: 20px;
            background: #fafafa;
            border-left: 2px solid #000;
        }}
        
        .calculator-title {{
            font-size: 0.9em;
            font-weight: 600;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        
        .calculator-input {{
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }}
        
        .calculator-input label {{
            font-size: 0.9em;
            color: #666;
        }}
        
        .calculator-input input {{
            padding: 8px 12px;
            border: 1px solid #ccc;
            background: white;
            font-size: 1em;
            width: 100px;
            transition: border 0.3s ease;
        }}
        
        .calculator-input input:focus {{
            outline: none;
            border-color: #000;
        }}
        
        .calculator-input button {{
            padding: 8px 20px;
            background: #000;
            color: #fff;
            border: none;
            cursor: pointer;
            font-size: 0.9em;
            transition: all 0.3s ease;
        }}
        
        .calculator-input button:hover {{
            background: #333;
        }}
        
        .calculator-result {{
            margin-top: 15px;
            padding: 15px;
            background: white;
            border: 1px solid #e0e0e0;
            font-size: 0.95em;
            display: none;
            animation: fadeIn 0.4s ease-in;
        }}
        
        .calculator-result.show {{
            display: block;
        }}
        
        .status-message {{
            padding: 15px 20px;
            margin-top: 20px;
            border-left: 3px solid #000;
            background: #fafafa;
            font-size: 0.95em;
        }}
        
        .status-message.good {{
            border-left-color: #4caf50;
        }}
        
        .status-message.warning {{
            border-left-color: #ff9800;
        }}
        
        .absent-section {{
            margin-top: 20px;
            padding: 15px;
            background: #f5f5f5;
        }}
        
        .absent-section h4 {{
            font-size: 0.85em;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
            color: #666;
        }}
        
        .absent-dates {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }}
        
        .date-tag {{
            padding: 4px 10px;
            background: white;
            border: 1px solid #ddd;
            font-size: 0.85em;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 80px;
            padding-top: 30px;
            border-top: 1px solid #e0e0e0;
            color: #999;
            font-size: 0.85em;
        }}
        
        @media print {{
            .calculator-section {{
                display: none;
            }}
        }}
        
        @media (max-width: 600px) {{
            .subject-stats {{
                grid-template-columns: 1fr;
            }}
            
            .overall-stats {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Attendance Report</h1>
            <p style="color: #999; font-size: 0.9em; margin-top: 10px;">Generated on {self.report_date}</p>
        </div>
        
        <div class="student-info">
            <h2>Student Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">Name</span>
                    <span class="info-value">{self.student_name if self.student_name else 'N/A'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">SAP ID</span>
                    <span class="info-value">{self.sap_id if self.sap_id else 'N/A'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Program</span>
                    <span class="info-value">{self.program if self.program else 'N/A'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Batch</span>
                    <span class="info-value">{self.batch if self.batch else 'N/A'}</span>
                </div>
            </div>
        </div>
        
        <div class="overall-stats">
            <div class="stat-box">
                <div class="stat-label">Total Lectures</div>
                <div class="stat-value">{total_lectures}</div>
            </div>
            <div class="stat-box">
                <div class="stat-label">Present</div>
                <div class="stat-value">{total_present}</div>
            </div>
            <div class="stat-box">
                <div class="stat-label">Absent</div>
                <div class="stat-value">{total_absent}</div>
            </div>
            <div class="stat-box">
                <div class="stat-label">Percentage</div>
                <div class="stat-value">{overall_percentage:.1f}%</div>
            </div>
        </div>
        
        <div class="subjects-section">
            <h2 class="section-title">Subject-wise Attendance</h2>
"""
        
        # Add subject cards
        subject_index = 0
        for subject, data in sorted_subjects:
            percentage = (data['present'] / data['total'] * 100) if data['total'] > 0 else 0
            subject_id = f"subject_{subject_index}"
            
            # Calculate what's needed for 80%
            current_p = data['present']
            current_t = data['total']
            
            if percentage >= 80:
                status_message = "You're in good standing."
                status_class = "good"
            else:
                needed = 0
                while True:
                    new_present = current_p + needed
                    new_total = current_t + needed
                    new_percentage = (new_present / new_total) * 100
                    if new_percentage >= 80:
                        break
                    needed += 1
                    if needed > 100:
                        break
                status_message = f"Attend {needed} consecutive lecture(s) to reach 80%."
                status_class = "warning"
            
            html_content += f"""
            <div class="subject-card" style="animation-delay: {subject_index * 0.1}s;">
                <div class="subject-header">
                    <div class="subject-name">{subject}</div>
                    <div class="subject-percentage">{percentage:.1f}%</div>
                </div>
                
                <div class="subject-stats">
                    <div class="stat-item">
                        <div class="stat-item-label">Total</div>
                        <div class="stat-item-value">{data['total']}</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-item-label">Present</div>
                        <div class="stat-item-value">{data['present']}</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-item-label">Absent</div>
                        <div class="stat-item-value">{data['absent']}</div>
                    </div>
                </div>
                
                <div class="status-message {status_class}">
                    {status_message}
                </div>
"""
            
            # Add absent dates if any
            if data['absent'] > 0:
                html_content += f"""
                <div class="absent-section">
                    <h4>Absent Dates</h4>
                    <div class="absent-dates">
"""
                for date in data['absent_dates']:
                    html_content += f'                        <span class="date-tag">{date}</span>\n'
                
                html_content += """                    </div>
                </div>
"""
            
            # Add calculator section
            html_content += f"""
                <div class="calculator-section">
                    <div class="calculator-title">Attendance Calculator</div>
                    <div class="calculator-input">
                        <label for="total_hours_{subject_id}">Total Hours in Subject:</label>
                        <input type="number" id="total_hours_{subject_id}" min="1" value="{data['total']}" />
                        <button onclick="calculateAttendance('{subject_id}', {data['present']}, {data['total']})">Calculate</button>
                    </div>
                    <div class="calculator-result" id="result_{subject_id}"></div>
                </div>
            </div>
"""
            subject_index += 1
        
        html_content += """        </div>
        
        <div class="footer">
            <p>This report is generated based on attendance data from the system.</p>
            <p>For discrepancies, contact administration.</p>
        </div>
    </div>
    
    <script>
        function calculateAttendance(subjectId, currentPresent, currentTotal) {
            const totalHoursInput = document.getElementById('total_hours_' + subjectId);
            const resultDiv = document.getElementById('result_' + subjectId);
            
            const totalHours = parseInt(totalHoursInput.value);
            
            if (!totalHours || totalHours < 1) {
                resultDiv.innerHTML = '<p style="color: #666;">Please enter a valid number of total hours.</p>';
                resultDiv.classList.add('show');
                return;
            }
            
            // Calculate current attendance percentage
            const currentPercentage = (currentPresent / currentTotal) * 100;
            
            // Calculate how many lectures can be missed while maintaining 80%
            let maxCanMiss = 0;
            let tempTotal = currentTotal;
            
            while (tempTotal < totalHours) {
                const newTotal = tempTotal + 1;
                const newPercentage = (currentPresent / newTotal) * 100;
                
                if (newPercentage >= 80) {
                    maxCanMiss++;
                    tempTotal++;
                } else {
                    break;
                }
            }
            
            // Calculate how many need to attend if below 80%
            let needToAttend = 0;
            if (currentPercentage < 80) {
                let tempPresent = currentPresent;
                let tempTotal2 = currentTotal;
                
                while (tempTotal2 < totalHours) {
                    tempPresent++;
                    tempTotal2++;
                    const newPercentage = (tempPresent / tempTotal2) * 100;
                    needToAttend++;
                    
                    if (newPercentage >= 80) {
                        break;
                    }
                }
            }
            
            let resultHTML = '<div style="padding: 10px;">';
            resultHTML += '<p style="margin-bottom: 10px;"><strong>Based on ' + totalHours + ' total hours:</strong></p>';
            
            if (currentPercentage >= 80 && maxCanMiss > 0) {
                resultHTML += '<p>You can miss up to <strong>' + maxCanMiss + ' lecture(s)</strong> and still maintain 80% attendance.</p>';
                resultHTML += '<p style="margin-top: 10px; color: #666;">Remaining lectures after that: ' + (totalHours - currentTotal - maxCanMiss) + '</p>';
            } else if (currentPercentage >= 80 && maxCanMiss === 0) {
                resultHTML += '<p>You cannot miss any more lectures to maintain 80% attendance.</p>';
                resultHTML += '<p style="margin-top: 10px; color: #666;">Remaining lectures: ' + (totalHours - currentTotal) + '</p>';
            } else {
                if (needToAttend > 0 && needToAttend <= (totalHours - currentTotal)) {
                    resultHTML += '<p>You need to attend <strong>' + needToAttend + ' consecutive lecture(s)</strong> to reach 80%.</p>';
                    const afterThat = totalHours - currentTotal - needToAttend;
                    if (afterThat > 0) {
                        resultHTML += '<p style="margin-top: 10px; color: #666;">After that, you\'ll have ' + afterThat + ' lecture(s) remaining.</p>';
                    }
                } else {
                    resultHTML += '<p style="color: #999;">With the remaining lectures, reaching 80% may not be possible.</p>';
                    resultHTML += '<p style="margin-top: 10px;">Current: ' + currentPercentage.toFixed(1) + '%</p>';
                }
            }
            
            resultHTML += '</div>';
            
            resultDiv.innerHTML = resultHTML;
            resultDiv.classList.add('show');
        }
    </script>
</body>
</html>"""
        
        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"\nHTML report generated: {output_file}")
        return output_file
    
    def export_to_csv(self, output_file='attendance_report.csv'):
        """Export attendance data to CSV"""
        if not self.attendance_data:
            print("No data to export.")
            return
        
        df = pd.DataFrame(self.attendance_data)
        df['subject_cleaned'] = df['course'].apply(self.clean_course_name)
        df.to_csv(output_file, index=False)
        print(f"Data exported to: {output_file}")
    
    def export_summary_to_csv(self, output_file='attendance_summary.csv'):
        """Export summary to CSV"""
        if not self.subjects:
            print("No summary data to export.")
            return
        
        summary_data = []
        for subject, data in self.subjects.items():
            percentage = (data['present'] / data['total'] * 100) if data['total'] > 0 else 0
            summary_data.append({
                'Subject': subject,
                'Total Lectures': data['total'],
                'Present': data['present'],
                'Absent': data['absent'],
                'Attendance %': f"{percentage:.2f}",
                'Status': 'Safe' if percentage >= 75 else 'Warning' if percentage >= 70 else 'Critical'
            })
        
        df = pd.DataFrame(summary_data)
        df = df.sort_values('Attendance %', ascending=False)
        df.to_csv(output_file, index=False)
        print(f"Summary exported to: {output_file}")


def main():
    print("\n" + "=" * 80)
    print("STUDENT ATTENDANCE CALCULATOR")
    print("=" * 80 + "\n")
    
    # Check if PDF path is provided
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = "ZSVKM_STUDENT_ATTENDANCE.pdf"
    
    # Check if file exists
    if not Path(pdf_path).exists():
        print(f"Error: File not found: {pdf_path}")
        print(f"\nUsage: python attendance_calculator.py [pdf_file]")
        sys.exit(1)
    
    # Create calculator instance
    calc = AttendanceCalculator(pdf_path)
    
    # Extract data
    if not calc.extract_data():
        sys.exit(1)
    
    # Calculate attendance
    calc.calculate_subject_attendance()
    
    # Generate HTML report
    html_file = calc.generate_html_report('attendance_report.html')
    
    # Also export CSV files
    calc.export_to_csv()
    calc.export_summary_to_csv()
    
    print("\nAnalysis complete!")
    print(f"Open '{html_file}' in your browser to view the report.\n")


if __name__ == "__main__":
    main()
