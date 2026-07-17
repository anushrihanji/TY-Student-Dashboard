"""app.py — Flask backend for TY Student Dashboard"""
from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from data_parser import LOADERS, DIVISION_PASSWORDS, DIVISION_SUBJECTS, DIVISION_ATT_SUBJECTS
import math

app = Flask(__name__)
app.secret_key = 'TY_DASHBOARD_SECRET_2024_XYZ_SECURE'

# ── AUTH ──────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    div = data.get('division','').upper()
    pwd = data.get('password','')
    if div not in DIVISION_PASSWORDS:
        return jsonify({'success':False,'message':'Invalid division'})
    if DIVISION_PASSWORDS[div] != pwd:
        return jsonify({'success':False,'message':'Incorrect password. Please try again.'})
    session['division'] = div
    return jsonify({'success':True})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

@app.route('/dashboard')
def dashboard():
    if 'division' not in session:
        return redirect(url_for('index'))
    return render_template('dashboard.html', division=session['division'])

# ── ATTENDANCE API ─────────────────────────────────────────────
@app.route('/api/attendance')
def api_attendance():
    if 'division' not in session:
        return jsonify({'error':'Unauthorized'}), 401
    div = session['division']
    month = request.args.get('month','')
    try:
        att = LOADERS[div]['attendance']()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    month_keys = list(att.keys())
    if not month or month not in att:
        month = month_keys[0] if month_keys else ''

    selected = att.get(month, {})
    all_students = []
    for batch_name, students in selected.items():
        for s in students:
            s['batch'] = batch_name
            all_students.append(s)

    defaulters = [s for s in all_students if s['attendance_pct'] < 60]
    low        = [s for s in all_students if 60 <= s['attendance_pct'] < 75]
    good       = [s for s in all_students if s['attendance_pct'] >= 75]

    return jsonify({
        'months': month_keys,
        'selected_month': month,
        'batches': sorted(set(s['batch'] for s in all_students)),
        'students': all_students,
        'defaulters': defaulters,
        'low': low,
        'good': good,
    })

# ── STUDENT ATTENDANCE DETAIL (for modal) ─────────────────────
@app.route('/api/student_att/<int:roll>')
def api_student_att(roll):
    if 'division' not in session:
        return jsonify({'error':'Unauthorized'}), 401
    div = session['division']
    try:
        att = LOADERS[div]['attendance']()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    month_keys = list(att.keys())
    feb_label = month_keys[0] if month_keys else ''
    mar_label = month_keys[1] if len(month_keys) > 1 else ''

    feb_student = mar_student = None
    for students in att.get(feb_label, {}).values():
        f = next((s for s in students if s['roll'] == roll), None)
        if f: feb_student = f; break
    for students in att.get(mar_label, {}).values():
        m = next((s for s in students if s['roll'] == roll), None)
        if m: mar_student = m; break

    feb_pct = feb_student['attendance_pct'] if feb_student else 0
    mar_pct = mar_student['attendance_pct'] if mar_student else 0
    ovr_pct = round((feb_pct + mar_pct) / 2, 2) if (feb_student or mar_student) else 0

    subjects = {}
    ref = feb_student or mar_student
    if ref and 'subjects' in ref:
        for subj, d in ref['subjects'].items():
            feb_att = d.get('feb_att', 0)
            feb_tot = d.get('feb_total', 0)
            mar_att = d.get('mar_att', 0)
            mar_tot = d.get('mar_total', 0)
            # If mar_student has separate data, use it
            if mar_student and 'subjects' in mar_student and subj in mar_student['subjects']:
                mar_att = mar_student['subjects'][subj].get('mar_att', mar_att)
                mar_tot = mar_student['subjects'][subj].get('mar_total', mar_tot)
            subjects[subj] = {
                'feb_att': feb_att, 'feb_total': feb_tot,
                'mar_att': mar_att, 'mar_total': mar_tot,
            }

    return jsonify({
        'roll': roll, 'feb_pct': feb_pct, 'mar_pct': mar_pct, 'ovr_pct': ovr_pct,
        'subjects': subjects,
    })

# ── MARKS API ─────────────────────────────────────────────────
@app.route('/api/marks')
def api_marks():
    if 'division' not in session:
        return jsonify({'error':'Unauthorized'}), 401
    div = session['division']
    subjects = DIVISION_SUBJECTS[div]
    try:
        students = LOADERS[div]['marks']()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    # AB lists (subject-wise)
    ab_ise1, ab_ise2 = [], []
    for s in students:
        for sub in subjects:
            sm = s['subjects'].get(sub, {})
            if sm.get('ISE1') == 'AB':
                ab_ise1.append({'roll':s['roll'],'grn':s['grn'],'name':s['name'],'subject':sub})
            if sm.get('ISE2') == 'AB':
                ab_ise2.append({'roll':s['roll'],'grn':s['grn'],'name':s['name'],'subject':sub})

    # Subject-wise avg < 16 and >= 16
    avg_lt16, avg_gte16 = [], []
    for s in students:
        for sub in subjects:
            v = s['subjects'].get(sub, {}).get('average')
            if isinstance(v, int):
                entry = {'roll':s['roll'],'grn':s['grn'],'name':s['name'],'subject':sub,'avg':v}
                if v < 16: avg_lt16.append(entry)
                else: avg_gte16.append(entry)

    # Pass/fail per subject (avg >= 16 = pass)
    pass_fail = {}
    total = len(students)
    for sub in subjects:
        passed = sum(1 for s in students if isinstance(s['subjects'].get(sub,{}).get('average'),int) and s['subjects'][sub]['average'] >= 16)
        failed = total - passed
        pass_fail[sub] = {
            'pass': passed, 'fail': failed,
            'pass_pct': round(passed/total*100,1) if total else 0,
            'fail_pct': round(failed/total*100,1) if total else 0,
        }

    return jsonify({
        'students': students,
        'subjects': subjects,
        'ab_ise1': ab_ise1,
        'ab_ise2': ab_ise2,
        'avg_lt16': avg_lt16,
        'avg_gte16': avg_gte16,
        'pass_fail': pass_fail,
    })

# ── STUDENT MARKS DETAIL ───────────────────────────────────────
@app.route('/api/student/<int:roll>')
def api_student(roll):
    if 'division' not in session:
        return jsonify({'error':'Unauthorized'}), 401
    div = session['division']
    subjects = DIVISION_SUBJECTS[div]
    try:
        students = LOADERS[div]['marks']()
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    student = next((s for s in students if s['roll'] == roll), None)
    if not student:
        return jsonify({'error':'Student not found'}), 404
    return jsonify({**student, 'subject_list': subjects})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
