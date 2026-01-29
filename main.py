import pyodbc
import getpass
import sys
import os


# ========= CONFIG =========
# Update this path if your database file has a different name/location.
ACCESS_DB_PATH = r"studenst.accdb"


def find_access_driver():
    """Find available Microsoft Access ODBC driver."""
    drivers = [d for d in pyodbc.drivers() if 'access' in d.lower() or 'mdb' in d.lower()]
    
    # Common driver names to try (in order of preference)
    preferred_drivers = [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.mdb)",
        "Driver do Microsoft Access (*.mdb)",  # Non-English systems
    ]
    
    # Check if any preferred driver is available
    for driver in preferred_drivers:
        if driver in drivers:
            return driver
    
    # If no preferred driver found, return first available Access driver
    if drivers:
        return drivers[0]
    
    return None


def get_connection():
    """
    Open a connection to the Access database using the ACE OLEDB provider.

    Requirements:
    - Microsoft Access Database Engine (ACE) 32/64-bit matching your Python.
    - `pyodbc` installed: pip install pyodbc
    """
    driver = find_access_driver()
    
    if not driver:
        print("\nERROR: Microsoft Access ODBC Driver not found!")
        print("\nPlease install Microsoft Access Database Engine:")
        print("Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        print("\nIMPORTANT: Install the version (32-bit or 64-bit) that matches your Python installation.")
        print("To check your Python version, run: python -c \"import platform; print(platform.architecture()[0])\"")
        sys.exit(1)
    
    # Convert relative path to absolute path
    db_path = os.path.abspath(ACCESS_DB_PATH)
    
    if not os.path.exists(db_path):
        print(f"\nWARNING: Database file not found at: {db_path}")
        print("Creating new database file...")
    
    conn_str = (
        rf"Driver={{{driver}}};"
        rf"Dbq={db_path};"
        r"Uid=Admin;"
        r"Pwd=;"
    )
    
    try:
        return pyodbc.connect(conn_str)
    except pyodbc.Error as e:
        print(f"\nERROR connecting to database: {e}")
        print(f"Driver used: {driver}")
        print(f"Database path: {db_path}")
        raise


def table_exists(cursor, table_name):
    """Check if a table exists in the database."""
    try:
        cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")
        return True
    except pyodbc.Error:
        return False


def init_db():
    """Create tables if they do not exist."""
    try:
        with get_connection() as conn:
            cur = conn.cursor()

            # Users table for Register/Login
            if not table_exists(cur, "Users"):
                cur.execute(
                    """
                    CREATE TABLE Users (
                        id AUTOINCREMENT PRIMARY KEY,
                        username TEXT(50) UNIQUE,
                        password TEXT(255)
                    )
                    """
                )
                print("Created Users table.")

            # Students table
            if not table_exists(cur, "Students"):
                cur.execute(
                    """
                    CREATE TABLE Students (
                        id AUTOINCREMENT PRIMARY KEY,
                        student_id TEXT(20) UNIQUE,
                        name TEXT(100),
                        score DOUBLE
                    )
                    """
                )
                print("Created Students table.")
            
            conn.commit()
    except Exception as e:
        print(f"Error initializing database: {e}")
        raise


# ========= AUTH =========
def register():
    username = input("Enter new username: ").strip()
    password = getpass.getpass("Enter new password: ").strip()
    if not username or not password:
        print("Username and password cannot be empty.")
        return

    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO Users (username, password) VALUES (?, ?)",
                (username, password),
            )
            conn.commit()
        print("User registered successfully.")
    except pyodbc.IntegrityError:
        print("Username already exists.")
    except Exception as e:
        print(f"Error during registration: {e}")


def login():
    username = input("Username: ").strip()
    password = getpass.getpass("Password: ").strip()

    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id FROM Users WHERE username = ? AND password = ?",
                (username, password),
            )
            row = cur.fetchone()
            if row:
                print(f"Welcome, {username}!")
                return True
            else:
                print("Invalid username or password.")
                return False
    except Exception as e:
        print(f"Error during login: {e}")
        return False


# ========= STUDENT OPERATIONS =========
def add_student():
    student_id = input("Student ID: ").strip()
    name = input("Name: ").strip()
    try:
        score = float(input("Score (0-100): ").strip())
    except ValueError:
        print("Invalid score.")
        return

    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO Students (student_id, name, score) VALUES (?, ?, ?)",
                (student_id, name, score),
            )
            conn.commit()
        print("Student added.")
    except pyodbc.IntegrityError:
        print("Student ID already exists.")
    except Exception as e:
        print(f"Error adding student: {e}")


def show_all_students():
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, student_id, name, score FROM Students")
            rows = cur.fetchall()
            if not rows:
                print("No students found.")
                return
            print(f"{'DB_ID':<6} {'StuID':<10} {'Name':<25} {'Score':>6}")
            print("-" * 55)
            for r in rows:
                print(f"{r.id:<6} {r.student_id:<10} {r.name:<25} {r.score:>6.2f}")
    except Exception as e:
        print(f"Error fetching students: {e}")


def search_student_by_id():
    student_id = input("Enter Student ID to search: ").strip()
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id, student_id, name, score FROM Students WHERE student_id = ?",
                (student_id,),
            )
            row = cur.fetchone()
            if row:
                print(f"Found: ID={row.student_id}, Name={row.name}, Score={row.score}")
            else:
                print("Student not found.")
    except Exception as e:
        print(f"Error searching student: {e}")


def search_students_by_name():
    name = input("Enter (part of) name to search: ").strip()
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id, student_id, name, score FROM Students WHERE name LIKE ?",
                (f"%{name}%",),
            )
            rows = cur.fetchall()
            if not rows:
                print("No matching students.")
                return
            for r in rows:
                print(f"ID={r.student_id}, Name={r.name}, Score={r.score}")
    except Exception as e:
        print(f"Error searching students: {e}")


def update_student():
    student_id = input("Enter Student ID to update: ").strip()
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id, student_id, name, score FROM Students WHERE student_id = ?",
                (student_id,),
            )
            row = cur.fetchone()
            if not row:
                print("Student not found.")
                return

            print(f"Current Name: {row.name}, Score: {row.score}")
            new_name = input("New name (leave blank to keep): ").strip()
            new_score_input = input("New score (leave blank to keep): ").strip()

            name_final = new_name if new_name else row.name
            score_final = row.score
            if new_score_input:
                try:
                    score_final = float(new_score_input)
                except ValueError:
                    print("Invalid score. Keeping old value.")

            cur.execute(
                "UPDATE Students SET name = ?, score = ? WHERE student_id = ?",
                (name_final, score_final, student_id),
            )
            conn.commit()
            print("Student updated.")
    except Exception as e:
        print(f"Error updating student: {e}")


def delete_student():
    student_id = input("Enter Student ID to delete: ").strip()
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "DELETE FROM Students WHERE student_id = ?",
                (student_id,),
            )
            if cur.rowcount == 0:
                print("Student not found.")
            else:
                conn.commit()
                print("Student deleted.")
    except Exception as e:
        print(f"Error deleting student: {e}")


def count_students():
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM Students")
            count = cur.fetchone()[0]
            print(f"Total students: {count}")
    except Exception as e:
        print(f"Error counting students: {e}")


def calculate_average_score():
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT AVG(score) FROM Students")
            avg = cur.fetchone()[0]
            if avg is None:
                print("No students to calculate average.")
            else:
                print(f"Average score: {avg:.2f}")
    except Exception as e:
        print(f"Error calculating average score: {e}")


def show_passed_students(pass_mark=50.0):
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT student_id, name, score FROM Students WHERE score >= ?", (pass_mark,))
            rows = cur.fetchall()
            if not rows:
                print("No passed students.")
                return
            print("Passed students:")
            for r in rows:
                print(f"ID={r.student_id}, Name={r.name}, Score={r.score}")
    except Exception as e:
        print(f"Error fetching passed students: {e}")


def show_failed_students(pass_mark=50.0):
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT student_id, name, score FROM Students WHERE score < ?", (pass_mark,))
            rows = cur.fetchall()
            if not rows:
                print("No failed students.")
                return
            print("Failed students:")
            for r in rows:
                print(f"ID={r.student_id}, Name={r.name}, Score={r.score}")
    except Exception as e:
        print(f"Error fetching failed students: {e}")


# ========= MENUS =========
def print_main_menu():
    print("\n=== Student Management System ===")
    print("1. Register")
    print("2. Login")
    print("3. Exit")


def print_student_menu():
    print("\n=== Student Operations ===")
    print("1. Add student")
    print("2. Show all students")
    print("3. Search student by ID")
    print("4. Search students by name")
    print("5. Update student")
    print("6. Delete student")
    print("7. Count students")
    print("8. Calculate average score")
    print("9. Show passed students")
    print("10. Show failed students")
    print("11. Exit")


def student_menu_loop():
    while True:
        print_student_menu()
        choice = input("Choose an option (1-11): ").strip()

        if choice == "1":
            add_student()
        elif choice == "2":
            show_all_students()
        elif choice == "3":
            search_student_by_id()
        elif choice == "4":
            search_students_by_name()
        elif choice == "5":
            update_student()
        elif choice == "6":
            delete_student()
        elif choice == "7":
            count_students()
        elif choice == "8":
            calculate_average_score()
        elif choice == "9":
            show_passed_students()
        elif choice == "10":
            show_failed_students()
        elif choice == "11":
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Try again.")


def main():
    init_db()
    while True:
        print_main_menu()
        choice = input("Choose an option (1-3): ").strip()

        if choice == "1":
            register()
        elif choice == "2":
            if login():
                student_menu_loop()
        elif choice == "3":
            print("Exiting program.")
            sys.exit(0)
        else:
            print("Invalid choice. Try again.")


if __name__ == "__main__":
    main()
