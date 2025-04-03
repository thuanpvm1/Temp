import csv
import logging
from ldap3 import Server, Connection, ALL, SUBTREE, MODIFY_REPLACE

# Configure logging
logging.basicConfig(
    filename='student_ad_update.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# AD connection settings
AD_SERVER = 'ldap://10.10.51.12'
AD_USER = 'khaisangcorp\\admin.thuanphan'
AD_PASSWORD = 'Kh@1Sang123$%^'
BASE_DN = 'DC=khaisangcorp,DC=vn'

def connect_to_ad():
    try:
        server = Server(AD_SERVER, get_info=ALL)
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        logger.info("Successfully connected to AD server")
        return conn
    except Exception as e:
        logger.error(f"Failed to connect to AD server: {str(e)}")
        raise

def update_std_common(conn, student):
    """Update basic AD attributes based on pupil email address"""
    search_filter = f"(mail={student['Pupil Email Address']})"
    attributes = ['distinguishedName']
    
    conn.search(BASE_DN, search_filter, search_scope=SUBTREE, attributes=attributes)
    
    if len(conn.entries) > 0:
        dn = conn.entries[0].distinguishedName.value
        changes = {
            'employeeID': [(MODIFY_REPLACE, [student['School Code']])],
            'employeeNumber': [(MODIFY_REPLACE, [student['School Id']])],
            'title': [(MODIFY_REPLACE, ['Student'])]
        }
        try:
            conn.modify(dn, changes)
            logger.info(f"Updated basic attributes for {student['Pupil Email Address']}")
        except Exception as e:
            logger.error(f"Failed to update basic attributes for {student['Pupil Email Address']}: {str(e)}")
    else:
        logger.warning(f"No AD account found for {student['Pupil Email Address']}")

def update_company_acc(conn, student):
    """Update boarding house specific attributes"""
    search_filter = f"(mail={student['Pupil Email Address']})"
    conn.search(BASE_DN, search_filter, search_scope=SUBTREE, attributes=['distinguishedName'])
    
    if len(conn.entries) > 0:
        dn = conn.entries[0].distinguishedName.value
        changes = {}
        
        if student['Boarding House'] == "Nam Long":
            changes = {
                'physicalDeliveryOfficeName': [(MODIFY_REPLACE, ['EMASI NAM LONG OFFICE'])],
                'company': [(MODIFY_REPLACE, ['EMASI NAM LONG SCHOOL'])]
            }
        elif student['Boarding House'] == "Van Phuc":
            changes = {
                'physicalDeliveryOfficeName': [(MODIFY_REPLACE, ['EMASI VAN PHUC OFFICE'])],
                'company': [(MODIFY_REPLACE, ['EMASI VAN PHUC SCHOOL'])]
            }
        
        if changes:
            try:
                conn.modify(dn, changes)
                logger.info(f"Updated boarding house attributes for {student['Pupil Email Address']}")
            except Exception as e:
                logger.error(f"Failed to update boarding house attributes for {student['Pupil Email Address']}: {str(e)}")

def update_dept_des_acc(conn, student):
    """Update department and description attributes"""
    search_filter = f"(mail={student['Pupil Email Address']})"
    conn.search(BASE_DN, search_filter, search_scope=SUBTREE, attributes=['distinguishedName'])
    
    if len(conn.entries) > 0:
        dn = conn.entries[0].distinguishedName.value
        changes = {}
        
        # Department logic
        year_code = student['Year Code']
        if year_code.startswith(('NLY', 'VPY')):
            year_num = year_code[3:]
            changes['department'] = [(MODIFY_REPLACE, [f'Year {year_num}'])]
        else:
            changes['department'] = [(MODIFY_REPLACE, [year_code])]
            
        # Description logic
        form = student['Form']
        if form.startswith(('NL-', 'VP-')):
            form_suffix = form[3:]
            changes['description'] = [(MODIFY_REPLACE, [f'Year 24-25 HR {form_suffix}'])]
            
        try:
            conn.modify(dn, changes)
            logger.info(f"Updated department/description for {student['Pupil Email Address']}")
        except Exception as e:
            logger.error(f"Failed to update department/description for {student['Pupil Email Address']}: {str(e)}")

def statistic_student_acc(conn, csv_file):
    """Compare student counts between CSV and AD"""
    csv_students = {'Nam Long': [], 'Van Phuc': []}
    ad_students = {'Nam Long': [], 'Van Phuc': []}
    
    # Count from CSV
    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row['Boarding House'] in csv_students:
                csv_students[row['Boarding House']].append(row['Pupil Email Address'])
    
    # Count from AD
    for house in ['Nam Long', 'Van Phuc']:
        if house == 'Nam Long':
            search_filter = '(company=EMASI NAM LONG SCHOOL)'
        else:
            search_filter = '(company=EMASI VAN PHUC SCHOOL)'
            
        conn.search(BASE_DN, search_filter, search_scope=SUBTREE, attributes=['mail'])
        for entry in conn.entries:
            ad_students[house].append(entry.mail.value)
    
    # Log results
    for house in ['Nam Long', 'Van Phuc']:
        csv_count = len(csv_students[house])
        ad_count = len(ad_students[house])
        
        logger.info(f"{house} - CSV student count: {csv_count}")
        logger.info(f"{house} - AD student count: {ad_count}")
        
        missing_in_ad = set(csv_students[house]) - set(ad_students[house])
        missing_in_csv = set(ad_students[house]) - set(csv_students[house])
        
        for email in missing_in_ad:
            logger.warning(f"{house} - Student in CSV but not in AD: {email}")
        for email in missing_in_csv:
            logger.warning(f"{house} - Student in AD but not in CSV: {email}")

def main(csv_file):
    try:
        conn = connect_to_ad()
        
        with open(csv_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for student in reader:
                update_std_common(conn, student)
                update_company_acc(conn, student)
                update_dept_des_acc(conn, student)
        
        statistic_student_acc(conn, csv_file)
        
        conn.unbind()
        logger.info("Script completed successfully")
        
    except Exception as e:
        logger.error(f"Script failed: {str(e)}")

if __name__ == "__main__":
    #csv_file_path = "C:/Users/ABC/PythonProjects/AI/students/ENL_Students.csv"
    main('C:/Users/ABC/PythonProjects/AI/students/ENL_Students.csv')  # Replace with your CSV file path
