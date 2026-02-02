import random
from collections import deque
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook


from config import CONFIG
from generators import *

# =========================
# ID REGISTRY
# =========================
class IDRegistry:
    """Registry to store and retrieve IDs for foreign key relationships"""
    
    def __init__(self, max_size):
        self.store = {}
        self.max_size = max_size

    def register(self, table, id_value):
        """Register an ID for a table"""
        if table not in self.store:
            self.store[table] = deque(maxlen=self.max_size)
        self.store[table].append(id_value)

    def get(self, table):
        """Get a random ID from a table"""
        if table not in self.store or not self.store[table]:
            raise ValueError(f"No IDs registered for {table}")
        return random.choice(self.store[table])
    
    def get_all(self, table):
        """Get all IDs from a table"""
        if table not in self.store:
            return []
        return list(self.store[table])
    
    def get_filtered(self, table, filter_fn):
        """Get IDs from a table that match a filter function"""
        all_ids = self.get_all(table)
        filtered = [id_val for id_val in all_ids if filter_fn(id_val)]
        if not filtered:
            raise ValueError(f"No IDs match filter for {table}")
        return random.choice(filtered)


# =========================
# EXCEL WRITER
# =========================
EXCEL_FILE_NAME = "Hospital_ERP_Data.xlsx"

def write_excel(table, rows, folder):
    """Write each table as a separate sheet in ONE Excel file"""
    if not rows:
        print(f"⚠️  Warning: No data generated for {table}")
        return

    Path(folder).mkdir(exist_ok=True)
    file_path = Path(folder) / EXCEL_FILE_NAME

    df = pd.DataFrame(rows)

    if file_path.exists():
        # Append as new sheet
        with pd.ExcelWriter(
            file_path,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            df.to_excel(writer, sheet_name=table, index=False)
    else:
        # Create new Excel file
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=table, index=False)

    print(f"✅ Written {len(rows):,} records to sheet '{table}'")


# =========================
# MAIN
# =========================
def main():
    print("=" * 60)
    print("🏥 Hospital ERP Data Generator")
    print("=" * 60)
    
    folder = CONFIG["output_folder"]
    registry = IDRegistry(CONFIG["max_active_ids"])

    # Date window
    start_date, end_date = get_date_range(
        CONFIG["start_date"],
        CONFIG["duration_years"]
    )
    print(f"📅 Date range: {start_date} to {end_date}")
    print()

    # ID counters
    state_id = 1
    city_id = 1
    branch_id = 1
    dept_id = 1
    spec_id = 1
    doc_id = 1
    room_type_id = 1
    room_id = 1
    patient_id = 1
    visit_id = 1
    billing_id = 1
    appointment_id = 1
    treatment_id = 1
    pharmacy_item_id = 1
    pharmacy_billing_detail_id = 1
    radiology_service_id = 1
    radiology_billing_detail_id = 1
    feedback_id = 1

    # -----------------------
    # MASTER DATA - GEOGRAPHY
    # -----------------------
    print("🌍 Generating Geography Data...")
    
    states = []
    for _ in range(CONFIG["states"]):
        states.append(gen_state(state_id))
        registry.register("state", state_id)
        state_id += 1
    write_excel("state", states, folder)

    cities = []
    for _ in range(CONFIG["cities"]):
        cities.append(gen_city(city_id, registry.get("state")))
        registry.register("city", city_id)
        city_id += 1
    write_excel("city", cities, folder)

    # -----------------------
    # MASTER DATA - HOSPITAL STRUCTURE
    # -----------------------
    print("\n🏥 Generating Hospital Structure...")
    
    branches = []
    for _ in range(CONFIG["hospital_branches"]):
        branches.append(gen_hospital_branch(branch_id, registry.get("city")))
        registry.register("hospital_branch", branch_id)
        branch_id += 1
    write_excel("hospital_branch", branches, folder)

    # ✅ FIXED: Generate departments FIRST
    departments = []
    for b in registry.get_all("hospital_branch"):
        for _ in range(CONFIG["departments_per_branch"]):
            departments.append(gen_department(dept_id, b))
            registry.register("department", dept_id)
            dept_id += 1
    write_excel("department", departments, folder)

    # ✅ FIXED: Generate specializations AFTER departments
    specializations = []
    specialization_lookup = {}
    for dept in registry.get_all("department"):
        for _ in range(CONFIG["specializations"] // len(registry.get_all("department")) + 1):
            specialization = gen_specialization(spec_id, dept)
            specializations.append(specialization)
            registry.register("specialization", spec_id)
            specialization_lookup[spec_id] = specialization
            spec_id += 1
    write_excel("specialization", specializations, folder)  # ✅ FIXED: Added write

    # -----------------------
    # MASTER DATA - DOCTORS
    # -----------------------
    print("\n👨‍⚕️ Generating Doctors...")
    
    doctors = []
    department_lookup = {d["id"]: d for d in departments}
    
    for _ in range(CONFIG["doctors"]):
        dept_id = registry.get("department")
        dept_branch = department_lookup[dept_id]["hospital_branch_id"]
        
        # ✅ FIXED: Proper indentation
        spec_id = registry.get_filtered(
            "specialization",
            lambda s: specialization_lookup[s]["department_id"] == dept_id
        )
        doctor = gen_doctor(
            doc_id,
            dept_id,
            spec_id,
            dept_branch
        )
        doctors.append(doctor)
        registry.register("doctor", doc_id)
        doc_id += 1

    # Assign exactly ONE HOD per department (FIXED)
    dept_to_doctors = {}
    for doctor in doctors:
        dept = doctor["department_id"]
        if dept not in dept_to_doctors:
            dept_to_doctors[dept] = []
        dept_to_doctors[dept].append(doctor)
    
    for dept, dept_docs in dept_to_doctors.items():
        if dept_docs:
            # Randomly select one doctor to be HOD
            hod_index = random.randint(0, len(dept_docs) - 1)
            dept_docs[hod_index]["is_hod"] = True

    write_excel("doctor", doctors, folder)

    # -----------------------
    # MASTER DATA - ROOMS
    # -----------------------
    print("\n🛏️  Generating Rooms...")
    
    room_types = []
    for _ in range(CONFIG["room_types"]):
        room_types.append(gen_room_type(room_type_id))
        registry.register("room_type", room_type_id)
        room_type_id += 1
    write_excel("room_type", room_types, folder)

    rooms = []
    for b in registry.get_all("hospital_branch"):
        for _ in range(CONFIG["rooms_per_branch"]):
            rooms.append(
                gen_room(
                    room_id,
                    registry.get("room_type"),
                    b
                )
            )
            registry.register("room", room_id)
            room_id += 1
    write_excel("room", rooms, folder)

    # -----------------------
    # MASTER DATA - TREATMENTS
    # -----------------------
    print("\n💊 Generating Treatments...")
    
    treatments = []
    treatment_lookup = {}
    
    for dept in registry.get_all("department"):
        for _ in range(CONFIG["treatments"] // len(registry.get_all("department")) + 1):
            treatment = gen_treatment(treatment_id, dept)
            treatments.append(treatment)
            registry.register("treatment", treatment_id)
            treatment_lookup[treatment_id] = treatment
            treatment_id += 1
    write_excel("treatment", treatments, folder)

    # -----------------------
    # MASTER DATA - PHARMACY ITEMS
    # -----------------------
    print("\n💊 Generating Pharmacy Items...")
    
    pharmacy_items = []
    for _ in range(CONFIG["pharmacy_items"]):
        pharmacy_items.append(gen_pharmacy_item(pharmacy_item_id))
        registry.register("pharmacy_item", pharmacy_item_id)
        pharmacy_item_id += 1
    write_excel("pharmacy_items", pharmacy_items, folder)

    # -----------------------
    # MASTER DATA - RADIOLOGY SERVICES
    # -----------------------
    print("\n🔬 Generating Radiology Services...")
    
    radiology_services = []
    for _ in range(CONFIG["radiology_services"]):
        radiology_services.append(gen_radiology_service(radiology_service_id))
        registry.register("radiology_service", radiology_service_id)
        radiology_service_id += 1
    write_excel("radiology_service", radiology_services, folder)

    # -----------------------
    # STATIC REFERENCE DATA
    # -----------------------
    print("\n📋 Generating Reference Data...")
    
    visit_types = gen_visit_types()
    for vt in visit_types:
        registry.register("visit_type", vt["id"])
    write_excel("visit_type", visit_types, folder)

    billing_types = gen_billing_types()
    for bt in billing_types:
        registry.register("billing_type", bt["id"])
    write_excel("billing_type", billing_types, folder)

    feedback_types = gen_feedback_types()
    for ft in feedback_types:
        registry.register("feedback_type", ft["id"])
    write_excel("feedback_type", feedback_types, folder)

    # -----------------------
    # TRANSACTIONAL DATA - PATIENTS
    # -----------------------
    print("\n👥 Generating Patients...")
    
    patients = []
    patient_data = []  # Store patient info for later use
    
    for _ in range(CONFIG["patients"]):
        patient = gen_patient(patient_id, start_date, end_date)
        patients.append(patient)
        registry.register("patient", patient_id)
        
        # Store patient info for visit generation
        patient_data.append({
            "id": patient_id,
            "created_date": patient["created_date"]
        })
        
        patient_id += 1
    
    write_excel("patient", patients, folder)

    # -----------------------
    # TRANSACTIONAL DATA - APPOINTMENTS, VISITS, BILLING
    # -----------------------
    print("\n📅 Generating Appointments, Visits, and Billing...")
    
    appointments = []
    visits = []
    billings = []
    pharmacy_billing_details = []
    radiology_billing_details = []
    
    # Create a lookup for rooms by branch
    rooms_by_branch = {}
    for room in rooms:
        branch = room["hospital_branch_id"]
        if branch not in rooms_by_branch:
            rooms_by_branch[branch] = []
        rooms_by_branch[branch].append(room)
    
    total_visits = 0
    
    # Generate multiple visits per patient (FIXED)
    for patient_info in patient_data:
        p_id = patient_info["id"]
        p_created = patient_info["created_date"]
        
        # Each patient has 1-5 visits
        num_visits = random.randint(
            CONFIG["visits_per_patient_min"],
            CONFIG["visits_per_patient_max"]
        )
        
        for visit_num in range(num_visits):
            # Pick visit type with realistic distribution
            visit_type = random.choices(
                [1, 2, 3],  # OP, IP, DC
                weights=[0.7, 0.2, 0.1],  # 70% OP, 20% IP, 10% DC
                k=1
            )[0]
            
            # Pick doctor and branch
            doctor_id = registry.get("doctor")
            branch_id = registry.get("hospital_branch")
            
            # Create appointment ONLY for OP visits
            appt_id = None
            if visit_type == 1:  # OP
                appointments.append(
                    gen_appointment(
                        appointment_id,
                        p_id,
                        doctor_id,
                        branch_id,
                        p_created,
                        start_date,
                        end_date
                    )
                )
                appt_id = appointment_id
                appointment_id += 1
            
            # Assign room ONLY for IP/DC and from the SAME branch (FIXED)
            room = None
            if visit_type in (2, 3):  # IP, DC
                # Get available rooms from the same branch
                branch_rooms = rooms_by_branch.get(branch_id, [])
                available_rooms = [r for r in branch_rooms if r["status"] == "A"]
                
                if available_rooms:
                    room_data = random.choice(available_rooms)
                    room = room_data["id"]
                # If no rooms available, leave as None (realistic scenario)
            
            # Create visit
            # Get the department id for the selected doctor
            doctor_dept_id = next(d for d in doctors if d["id"] == doctor_id)["department_id"]
            doctor_dept_id= department_lookup[doctor_dept_id]

            # Filter treatments by department ID (was incorrectly comparing to department dict)
            treatment_id = (
                registry.get_filtered(
                    "treatment",
                    lambda t: treatment_lookup[t]["department_id"] == doctor_dept_id
                )
                if random.random() > 0.3
                else None
            )
            visits.append(
                gen_patient_visit(
                    visit_id,
                    p_id,
                    doctor_id,
                    branch_id,
                    visit_type,
                    appt_id,
                    treatment_id,
                    room,
                    p_created,
                    start_date,
                    end_date
                )
            )
            
            # Create billing records (1-3 per visit)
            num_bills = random.randint(1, 3)
            for _ in range(num_bills):
                billing_type_id = registry.get("billing_type")
                
                billings.append(
                    gen_billing(
                        billing_id,
                        visit_id,
                        branch_id,
                        billing_type_id
                    )
                )
                
                # Generate billing details for pharmacy and radiology
                if billing_type_id == 1:  # Pharmacy
                    if random.random() < CONFIG["pharmacy_billing_detail_probability"]:
                        num_items = random.randint(
                            CONFIG["pharmacy_items_per_bill_min"],
                            CONFIG["pharmacy_items_per_bill_max"]
                        )
                        for _ in range(num_items):
                            pharmacy_billing_details.append(
                                gen_pharmacy_billing_detail(
                                    pharmacy_billing_detail_id,
                                    billing_id,
                                    registry.get("pharmacy_item")
                                )
                            )
                            pharmacy_billing_detail_id += 1
                
                elif billing_type_id == 2:  # Radiology
                    if random.random() < CONFIG["radiology_billing_detail_probability"]:
                        num_items = random.randint(
                            CONFIG["radiology_items_per_bill_min"],
                            CONFIG["radiology_items_per_bill_max"]
                        )
                        for _ in range(num_items):
                            radiology_billing_details.append(
                                gen_radiology_billing_detail(
                                    radiology_billing_detail_id,
                                    billing_id,
                                    registry.get("radiology_service")
                                )
                            )
                            radiology_billing_detail_id += 1
                
                billing_id += 1
            
            visit_id += 1
            total_visits += 1

    write_excel("appointment", appointments, folder)
    write_excel("patient_visit", visits, folder)
    write_excel("billing", billings, folder)
    write_excel("pharmacy_billing_detail", pharmacy_billing_details, folder)
    write_excel("radiology_billing_detail", radiology_billing_details, folder)

    # -----------------------
    # TRANSACTIONAL DATA - FEEDBACK
    # -----------------------
    print("\n💬 Generating Feedback...")
    
    feedbacks = []
    
    # 30% of patients leave feedback
    for patient_info in patient_data:
        if random.random() < CONFIG["feedback_probability"]:
            feedbacks.append(
                gen_feedback(
                    feedback_id,
                    patient_info["id"],
                    registry.get("feedback_type")
                )
            )
            feedback_id += 1
    
    write_excel("feedback", feedbacks, folder)

    # -----------------------
    # SUMMARY
    # -----------------------
    print("\n" + "=" * 60)
    print("✅ Hospital ERP dataset generated successfully!")
    print("=" * 60)
    print(f"📊 Summary:")
    print(f"   States: {len(states):,}")
    print(f"   Cities: {len(cities):,}")
    print(f"   Hospital Branches: {len(branches):,}")
    print(f"   Departments: {len(departments):,}")
    print(f"   Specializations: {len(specializations):,}")
    print(f"   Doctors: {len(doctors):,}")
    print(f"   Room Types: {len(room_types):,}")
    print(f"   Rooms: {len(rooms):,}")
    print(f"   Treatments: {len(treatments):,}")
    print(f"   Pharmacy Items: {len(pharmacy_items):,}")
    print(f"   Radiology Services: {len(radiology_services):,}")
    print(f"   Patients: {len(patients):,}")
    print(f"   Appointments: {len(appointments):,}")
    print(f"   Patient Visits: {len(visits):,}")
    print(f"   Billing Records: {len(billings):,}")
    print(f"   Pharmacy Billing Details: {len(pharmacy_billing_details):,}")
    print(f"   Radiology Billing Details: {len(radiology_billing_details):,}")
    print(f"   Feedbacks: {len(feedbacks):,}")
    print(f"\n📁 Output folder: {folder}/")
    print("=" * 60)


# =========================
# ENTRY
# =========================
if __name__ == "__main__":
    main()