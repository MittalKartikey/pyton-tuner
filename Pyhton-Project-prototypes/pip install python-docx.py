

from doc import Document
from doc.shared import Pt

# Create a new Document
doc = Document()

# Title
doc.add_heading('Riya Sharma', level=1).alignment = 1
doc.add_paragraph('[Phone Number] | [Email Address]').alignment = 1
doc.add_paragraph('[LinkedIn Profile] | [Address in Jaipur, Rajasthan]').alignment = 1

# Professional Summary
doc.add_heading('Professional Summary', level=2)
doc.add_paragraph(
    "Dedicated and results-driven education management professional with over 10 years of experience in school administration, curriculum development, "
    "and team leadership. Proven track record of enhancing educational standards, managing budgets, and fostering a positive learning environment. "
    "Adept at strategic planning, stakeholder engagement, and implementing innovative educational practices. Passionate about driving student success and teacher development in the vibrant educational landscape of Jaipur."
)

# Experience
doc.add_heading('Experience', level=2)

# Experience Entry 1
doc.add_heading('Principal', level=3)
doc.add_paragraph('Rajasthan International School, Jaipur')
doc.add_paragraph('July 2018 – Present')
doc.add_paragraph(
    "- Spearheaded the development and execution of a comprehensive school improvement plan, resulting in a 20% increase in student academic performance and a 15% boost in parent satisfaction scores."
)
doc.add_paragraph(
    "- Managed a team of 40+ teachers and administrative staff, providing leadership, support, and professional development opportunities."
)
doc.add_paragraph(
    "- Implemented a new digital learning initiative, integrating technology into the curriculum and improving student engagement by 25%."
)
doc.add_paragraph(
    "- Oversaw budgeting and financial planning for school operations, achieving a 10% reduction in operational costs while maintaining quality educational services."
)

# Experience Entry 2
doc.add_heading('Assistant School Director', level=3)
doc.add_paragraph('Jaipur Academy, Jaipur')
doc.add_paragraph('May 2013 – June 2018')
doc.add_paragraph(
    "- Assisted in the daily administration of the school, including staff management, student enrollment, and curriculum oversight."
)
doc.add_paragraph(
    "- Developed and introduced innovative programs for student enrichment and extracurricular activities, contributing to a more diverse and engaging learning environment."
)
doc.add_paragraph(
    "- Coordinated with parents, community leaders, and educational authorities to ensure alignment with educational standards and foster community involvement."
)
doc.add_paragraph(
    "- Conducted performance evaluations for teaching staff, identifying areas for improvement and facilitating targeted training workshops."
)

# Experience Entry 3
doc.add_heading('Educational Coordinator', level=3)
doc.add_paragraph('Shree Krishna School, Jaipur')
doc.add_paragraph('August 2008 – April 2013')
doc.add_paragraph(
    "- Designed and implemented curriculum enhancements, focusing on integrating interdisciplinary approaches and experiential learning."
)
doc.add_paragraph(
    "- Organized and managed school events, including parent-teacher meetings, educational fairs, and workshops, enhancing community engagement and support."
)
doc.add_paragraph(
    "- Analyzed student performance data to identify trends and areas for improvement, driving targeted interventions and support programs."
)

# Education
doc.add_heading('Education', level=2)

doc.add_heading('Master of Education (M.Ed.)', level=3)
doc.add_paragraph('University of Rajasthan, Jaipur')
doc.add_paragraph('2006 – 2008')

doc.add_heading('Bachelor of Education (B.Ed.)', level=3)
doc.add_paragraph('University of Rajasthan, Jaipur')
doc.add_paragraph('2003 – 2006')

doc.add_heading('Bachelor of Arts (B.A.) in English', level=3)
doc.add_paragraph('University of Rajasthan, Jaipur')
doc.add_paragraph('2000 – 2003')

# Certifications
doc.add_heading('Certifications', level=2)

doc.add_paragraph('- Certified School Administrator – National Association of School Administrators')
doc.add_paragraph('- Leadership in Education – Indian Institute of Management Bangalore (IIMB)')
doc.add_paragraph('- Advanced Certificate in Digital Education – National Institute of Educational Technology')

# Skills
doc.add_heading('Skills', level=2)

doc.add_paragraph('- Strategic Planning and Implementation')
doc.add_paragraph('- Budgeting and Financial Management')
doc.add_paragraph('- Curriculum Development and Improvement')
doc.add_paragraph('- Team Leadership and Staff Development')
doc.add_paragraph('- Parent and Community Engagement')
doc.add_paragraph('- Technology Integration in Education')
doc.add_paragraph('- Data Analysis and Performance Metrics')
doc.add_paragraph('- Conflict Resolution and Problem Solving')

# Professional Affiliations
doc.add_heading('Professional Affiliations', level=2)

doc.add_paragraph('- Member, Rajasthan Association of School Administrators')
doc.add_paragraph('- Member, National Association of Principals')
doc.add_paragraph('- Member, Jaipur Educational Leadership Network')

# Languages
doc.add_heading('Languages', level=2)

doc.add_paragraph('- English (Fluent)')
doc.add_paragraph('- Hindi (Fluent)')
doc.add_paragraph('- Rajasthani (Conversational)')

# References
doc.add_heading('References', level=2)
doc.add_paragraph('Available upon request.')

# Save the document
doc.save('Riya_Sharma_Resume.docx')
