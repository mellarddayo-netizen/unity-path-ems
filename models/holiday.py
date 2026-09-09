from extensions import db

class Holiday(db.Model):
    __tablename__ = "holidays"

    id = db.Column(db.Integer, primary_key=True)
    holiday_date = db.Column(db.Date, nullable=False, unique=True)
    name = db.Column(db.String(200), nullable=False)
    holiday_type = db.Column(db.String(50), nullable=False, default="Regular Holiday")
    # Multiplier for the first 8 hours when the employee works the holiday.
    work_multiplier = db.Column(db.Float, default=1.0)
    # Multiplier paid when the employee does not work. 0 means unpaid.
    unworked_multiplier = db.Column(db.Float, default=0.0)
    # OT multiplier applied after the holiday work multiplier.
    ot_multiplier = db.Column(db.Float, default=1.30)
    active = db.Column(db.Boolean, default=True)

    def __repr__(self):
        return f"<Holiday {self.holiday_date} {self.name}>"
