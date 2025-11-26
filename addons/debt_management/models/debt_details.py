import io
import math
import xlsxwriter
import base64
from odoo import api, fields, models
from dateutil.relativedelta import relativedelta
from datetime import date, timedelta


class DebtDetails(models.Model):
    _name = 'debt.details'
    _rec_name = 'loan_no'

    loan_no = fields.Char(string="Loan Number", required=True)
    sanctioned_amount = fields.Float(string="Sanctioned Amount", required=True)
    actual_amount = fields.Float(string="Actual Loan Amount", default=0)
    principal_amount = fields.Float(string="Principal Amount", default=0)
    remaining_amount = fields.Float(string="Remaining Use Amount", compute='_compute_remain_amount', store=True)
    loan_type = fields.Selection(
        [
            ('personal', 'Personal Loan'),
            ('debt_consolidation', 'Debt Consolidation Loan'),
            ('mortgage', 'Mortgage'),
            ('home_equity', 'Home Equity Loan'),
            ('student', 'Student Loan'),
            ('auto', 'Auto Loan'),
            ('small_business', 'Small Business Loan'),
            ('credit_builder', 'Credit Builder Loan'),
            ('payday', 'Payday Loan')
        ],
        string='Loan Type',
        required=True
    )
    loan_bank = fields.Many2one('res.bank', string="Bank Name")
    starting_date = fields.Date(string="Loan Sanctioned Date")
    loan_tenor = fields.Integer(string="Loan Tenure (Months)", required=True, default=0)
    first_emi = fields.Date(string="First EMI Date")
    emi_date = fields.Date(string="Next EMI Date", compute='_compute_emi_date', store=True)
    last_date = fields.Date(string="Last EMI Date", compute='_compute_last_month', store=True)
    interest_rate = fields.Float(string="Interest Rate (%)", required=True, default=9.0)

    # Emi Fields
    emi_amount = fields.Float(string="EMI Amount", compute='_compute_emi_amount', store=True)
    emi_remaining = fields.Integer(string="EMIs Remaining", compute='_compute_emi_remaining', store=True)
    emi_paid = fields.Integer(string="EMIs Paid", compute='_compute_emi_paid', store=True)
    total_interest = fields.Float(string="Total Interest Payable", compute='_compute_total_interest', store=True)
    total_debt = fields.Float(string="Total Debt", compute='_compute_total_debt', store=True)
    remaining_debt = fields.Float(string="Remaining Debt", compute='_compute_remaining_debt', store=True)
    # Advance Payment section
    advance_pay = fields.Boolean(string="Advance Pay", default=False)
    penalty_applicable = fields.Boolean(string="Penalty Applicable", default=False)
    penalty_percentage = fields.Float(string="Penalty Percentage (%)", default=0.0)
    gst_percentage = fields.Float(string="GST Percentage (%)", default=0.0)
    advance_type = fields.Selection(
        [('partial', 'Partial'), ('full', 'Full')],
        string='Advance Type',
    )
    reduction_type = fields.Selection(
        [('tenor_reduction', 'Tenor Reduction'), ('emi_reduction', 'Emi Reduction')],
        string='Reduction Type',
    )
    statusbar = fields.Selection([('in_progress', 'In Progress'), ('completed', 'Completed')], string='Status',
                                 default="in_progress", required=True)
    advance_amount = fields.Float(string="Advance Payment Amount", default=0.0)
    document = fields.Binary(string="Sanctioned Document")
    total_payable = fields.Float(string='Total Payable', default=0.0, compute='_compute_totals')
    total_advance_payment = fields.Float(string="Total Advance Payments", compute='_compute_total_advance_payment',
                                         store=True)

    receipt = fields.Binary()
    # Reminder days for email notification
    reminder_days = fields.Integer(string="Reminder Days",
                                   help="Number of days before EMI due date to send a reminder.", default=7)
    email = fields.Char(string="Enter your email")
    debt_paid = fields.Float(string="Debt Paid", compute='_compute_debt_paid', store=True)

    _sql_constraints = [
        ('unique_loan_no', 'UNIQUE(loan_no)', 'The loan number must be unique!')
    ]

    ##### Compute Methods #########
    @api.depends('emi_remaining', 'emi_amount', 'advance_amount')
    def _compute_debt_paid(self):
        for record in self:
            # Total paid includes EMIs that have been paid and any advance payments made
            total_paid = (record.loan_tenor - record.emi_remaining) * record.emi_amount + record.advance_amount
            record.debt_paid = total_paid

    @api.depends('penalty_applicable', 'advance_amount', 'penalty_percentage', 'advance_type', 'gst_percentage', )
    def _compute_totals(self):
        for record in self:
            record.total_payable = 0.0  # Initialize total_payable to avoid incorrect values
            if record.advance_pay:
                penalty = 0.0
                gst = 0.0
                if record.penalty_applicable:
                    if record.advance_type == 'partial':
                        penalty = (record.advance_amount * (record.penalty_percentage / 100))
                    elif record.advance_type == 'full':
                        penalty = (record.remaining_debt * (record.penalty_percentage / 100))
                    gst = (penalty * record.gst_percentage) / 100
                if record.advance_type == 'partial':
                    record.total_payable = record.advance_amount + penalty + gst
                elif record.advance_type == 'full':
                    record.total_payable = record.remaining_debt + penalty + gst

            else:
                record.total_payable = 0.0

    @api.depends('actual_amount', 'sanctioned_amount')
    def _compute_remain_amount(self):
        for record in self:
            record.remaining_amount = max(0, record.sanctioned_amount - record.actual_amount)

    @api.depends('principal_amount', 'interest_rate')
    def _compute_emi_amount(self):
        for record in self:
            # Initialize EMI amount to zero
            record.emi_amount = 0.0

            # Check if inputs are valid using Odoo's float_compare method for precision safety
            if record.loan_tenor > 0 and record.sanctioned_amount > 0 and record.interest_rate >= 0:
                monthly_interest_rate = record.interest_rate / (12 * 100)  # Monthly interest rate as a fraction

                try:
                    # EMI Base Calculation using the formula
                    emi_base = (record.principal_amount * monthly_interest_rate) / (
                            1 - (1 + monthly_interest_rate) ** -record.loan_tenor)

                    # Ensure EMI is not negative and round it to 2 decimal places
                    record.emi_amount = round(max(emi_base, 0), 2)
                    print(emi_base)
                except ZeroDivisionError:
                    print(f"ZeroDivisionError while calculating EMI for record {record.id}: "
                          "Loan tenor may be invalid or zero.")
                    record.emi_amount = 0.0
                except Exception as e:
                    print(f"Error calculating EMI for record {record.id}: {str(e)}")
                    record.emi_amount = 0.0
            else:
                # Set EMI to 0 if input parameters are invalid
                record.emi_amount = 0.0

    @api.depends('first_emi', 'loan_tenor')
    def _compute_emi_date(self):
        for record in self:
            if record.first_emi:
                # Get the current date
                current_date = fields.Date.today()
                # Convert first_emi to a date object
                first_emi_date = fields.Date.from_string(record.first_emi)
                # If the current date is before the first EMI date, set emi_date to first_emi
                if current_date < first_emi_date:
                    record.emi_date = first_emi_date
                else:
                    # Calculate the number of months since the first EMI
                    months_difference = (current_date.year - first_emi_date.year) * 12 + (
                            current_date.month - first_emi_date.month)

                    # Calculate the next EMI date based on loan tenor
                    next_emi_months = months_difference + 1  # Next EMI after the last one paid
                    next_emi_date = first_emi_date + relativedelta(months=next_emi_months)

                    # Ensure that we don't exceed the loan tenor
                    if next_emi_months <= record.loan_tenor:
                        record.emi_date = next_emi_date
                    else:
                        record.emi_date = False  # No more EMIs due if tenor is exceeded
            else:
                record.emi_date = False  # Reset if no first EMI date is set

    @api.depends('first_emi', 'loan_tenor')
    def _compute_last_month(self):
        for record in self:
            if record.first_emi and record.loan_tenor:
                first_emi_date = fields.Date.from_string(record.first_emi)
                # Calculate the last date after the given loan tenure (loan_tenor - 1 month)
                last_month_date = record.loan_tenor - 1
                record.last_date = first_emi_date + relativedelta(months=last_month_date)
            else:
                record.last_date = False

    @api.depends('first_emi', 'loan_tenor')
    def _compute_emi_remaining(self):
        for record in self:
            if record.first_emi and record.loan_tenor > 0:
                # Query the debt.emi.history model for the current loan
                emi_history = self.env['debt.emi.history'].search([('loan_id', '=', record.id)])

                # Count the number of EMIs already paid (or created).
                # You can modify the condition here based on how the status is set (e.g., 'paid' status)
                paid_emis = sum(1 for emi in emi_history if emi.payment_status == 'paid')

                # Calculate the remaining EMIs
                record.emi_remaining = max(0, record.loan_tenor - paid_emis)
            else:
                record.emi_remaining = record.loan_tenor

    @api.depends('emi_remaining')
    def _compute_emi_paid(self):
        for record in self:
            emi_count = self.env['debt.emi.history'].search_count([
                ('loan_id', '=', record.id),
                ('payment_status', '=', 'paid')
            ])
            record.emi_paid = emi_count

    @api.depends('actual_amount', 'interest_rate', 'loan_tenor')
    def _compute_total_debt(self):
        for record in self:
            if record.loan_tenor > 0 and record.principal_amount > 0:
                # Calculate monthly interest rate
                monthly_interest_rate = record.interest_rate / 100 / 12
                # Number of months for the loan tenor
                number_of_months = record.loan_tenor

                # Calculate EMI using the standard formula
                emi = (record.actual_amount * monthly_interest_rate * (1 + monthly_interest_rate) ** number_of_months) / \
                      ((1 + monthly_interest_rate) ** number_of_months - 1)

                # Total debt is EMI multiplied by the number of months
                record.total_debt = round(emi * number_of_months, 2)
            else:
                record.total_debt = 0.0

    @api.depends('actual_amount', 'total_debt')
    def _compute_total_interest(self):
        for record in self:
            if record.actual_amount and record.total_debt:
                record.total_interest = record.total_debt - record.actual_amount
            else:
                record.total_interest = False

    @api.depends('total_debt')
    def _compute_remaining_debt(self):
        for record in self:
            record.remaining_debt = max(0, record.total_debt - record.debt_paid - record.advance_amount)

    @api.onchange('actual_amount')
    def _onchange_actual_amount(self):
        """ Recalculate total debt, EMI amounts, and remaining debt when actual_amount is changed """
        self._compute_total_debt()  # Recalculate the total debt
        self._compute_remaining_debt()  # Recalculate remaining debt
        self._compute_emi_amount()  # Recalculate EMI amounts

    # Methods creating Emi Records
    @api.model
    def create(self, vals):
        # Create the loan record
        loan = super(DebtDetails, self).create(vals)
        loan.principal_amount = loan.actual_amount
        # Ensure EMI amount is calculated before creating EMI records
        if loan.loan_tenor > 0 and loan.first_emi and loan.emi_amount > 0:
            # Generate EMI records for the entire loan tenure up until today
            today = fields.Date.today()

            # Check if advance_type is 'full' and skip creating future EMIs
            if loan.advance_type == 'full':
                # For full advance, we do not create future EMI records
                loan.emi_remaining = 0
                loan.emi_date = False
            else:
                # Generate EMI records for each month in the loan tenure
                for month in range(loan.loan_tenor):
                    due_date = loan.first_emi + relativedelta(months=month)

                    # Only create EMI records up to today's date (skip future months)
                    if due_date <= today:
                        # Create EMI record in the debt.emi.history model
                        self.env['debt.emi.history'].create({
                            'loan_id': loan.id,
                            'due_date': due_date,
                            'payment_amount': loan.emi_amount,  # The EMI amount
                            'payment_status': 'paid',  # Initially marked as paid
                        })
                    else:
                        break

        return loan

    def action_done(self):
        for rec in self:
            # If the loan is marked as completed, create an EMI history record with the final closing balance
            if rec.advance_type == 'full':
                # Set the loan status to 'completed'
                rec.statusbar = 'completed'

                # Set the emi_date to False as the loan is completed
                rec.emi_date = False
                print(rec.total_payable)
                # For 'full' advance_type, create the last EMI record with the closing balance (total_payable)
                self.env['debt.emi.history'].create({
                    'loan_id': rec.id,
                    'due_date': fields.Date.today(),  # Use current date if emi_date is not set
                    'payment_amount': 0,
                    'advance_payment': rec.total_payable,
                    'remaining_debt': 0,  # Closing balance
                    'payment_status': 'paid',  # Mark as paid
                })
                rec.remaining_debt = 0.0
                rec.advance_amount = False
                rec.advance_type = False
            elif rec.advance_type == 'partial':
                rec.emi_paid += 1
                self.env['debt.emi.history'].create({
                    'loan_id': rec.id,
                    'due_date': fields.Date.today(),  # Use current date if emi_date is not set
                    'payment_amount': 0,
                    'advance_payment': rec.advance_amount,
                    'remaining_debt': rec.remaining_debt,
                    'payment_status': 'paid',  # Mark as paid
                })
                if rec.reduction_type == 'emi_reduction':
                    # When EMI reduction is chosen, reduce principal and recalculate EMI
                    initial_principal = rec.principal_amount
                    rec.principal_amount -= rec.advance_amount
                    rec.remaining_debt -= rec.advance_amount

                    # Ensure principal does not go negative
                    if rec.principal_amount < 0:
                        rec.principal_amount = 0

                    # Log the values before recalculating EMI
                    print(f"Principal Amount before reduction: {initial_principal}")
                    print(f"Emi Amount before reduction -----> {rec.emi_amount}")
                    print(f"Principal Amount after reduction -----> {rec.principal_amount}")

                    # Recalculate EMI only if the principal has changed
                    if rec.principal_amount != initial_principal:
                        self._compute_emi_amount()
                else:
                    rec.remaining_debt -= rec.total_payable
                    # Ensure remaining debt doesn't go negative
                    if rec.remaining_debt < 0:
                        rec.remaining_debt = 0

                    rec.emi_remaining = math.ceil(rec.remaining_debt / rec.emi_amount)

                rec.advance_amount = False
                rec.advance_type = False

    # Compute the total advance payment across all related EMI records
    def _compute_total_advance_payment(self):
        for record in self:
            # Calculate the total advance payments made for the loan
            total_advance_payment = sum(
                emi.advance_payment for emi in self.env['debt.emi.history'].search([('loan_id', '=', record.id)]))
            record.total_advance_payment = total_advance_payment

    @api.onchange('actual_amount')
    # Method to update the principal amount based on the total advance payments
    def _update_principal_amount(self):
        for record in self:
            record.principal_amount = record.actual_amount - record.total_advance_payment
            # Ensure the principal amount does not go below zero
            if record.principal_amount < 0:
                record.principal_amount = 0.0

    # Methods for Cron Jobs
    @api.model
    def update_emi_dates_daily(self):
        """
        This method updates the emi_date for all loans on a daily basis.
        If the next EMI is due, it creates a record in the debt.emi.history model.
        """
        # Get today's date
        current_date = fields.Date.today()

        # Find all loans that have a valid first_emi
        loans = self.search([('first_emi', '!=', False)])

        for loan in loans:
            first_emi_date = fields.Date.from_string(loan.first_emi)
            if current_date < first_emi_date:
                continue
            else:
                # Calculate the number of months since the first EMI
                months_difference = (current_date.year - first_emi_date.year) * 12 + (
                        current_date.month - first_emi_date.month)
                next_emi_months = months_difference + 1
                next_emi_date = first_emi_date + relativedelta(months=next_emi_months)

                if current_date == loan.emi_date:

                    # Ensure that we don't exceed the loan tenor
                    if next_emi_months <= loan.loan_tenor:
                        loan.emi_date = next_emi_date

                        # Check if the next EMI date matches today's date
                        if loan.emi_date == current_date and loan.advance_type != 'full':
                            # Create a new record in the debt.emi.history model
                            self.env['debt.emi.history'].create({
                                'loan_id': loan.id,
                                'due_date': current_date,
                                'payment_amount': loan.emi_amount,  # The EMI amount
                                'payment_status': 'paid',  # Initially marked as paid
                            })
                    else:
                        loan.emi_date = False

    def send_emi_reminder_email(self):
        today = date.today()
        # Iterate over all debt records and check if the reminder is due
        for record in self:
            if record.emi_date:
                reminder_date = record.emi_date - timedelta(days=record.reminder_days)
                print(reminder_date)
                if reminder_date == today:
                    # Send email logic here
                    template = self.env.ref('debt_management.email_template')
                    self.env['mail.template'].browse(template.id).send_mail(record.id, force_send=True)

    # View for Emi Records
    def action_view_emi(self):
        return {
            'type': 'ir.actions.act_window',
            'name': 'EMI Records',
            'res_model': 'debt.emi.history',
            'view_mode': 'tree,form',  # You can specify which views you want here
            'domain': [('loan_id', '=', self.id)],  # Only records related to this loan
            'context': {'default_loan_id': self.id},  # Set a default context, if necessary
        }

    ## Excel Report Generating Method
    def action_generate_emi_report(self):
        """
        This method generates the Excel report for the EMI history of the loan.
        """

        # Create a workbook in memory
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('EMI History')
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

        # Fetch EMI history records related to this loan
        emi_records = self.env['debt.emi.history'].search([('loan_id', '=', self.id)])

        # Write the header row
        worksheet.write(0, 0, 'Payment Date')
        worksheet.write(0, 1, 'Payment Amount')
        worksheet.write(0, 2, 'Advance Payment')
        worksheet.write(0, 3, 'Remaining Debt')
        worksheet.write(0, 4, 'Payment Status')
        worksheet.write(0, 5, 'Bank Name')
        worksheet.write(0, 6, 'Interest Rate (%)')
        worksheet.write(0, 7, 'Loan Type')

        # Fill in the data rows
        row = 1
        for emi in emi_records:
            worksheet.write(row, 0, emi.due_date.strftime('%Y-%m-%d'), date_format)
            worksheet.write(row, 1, emi.payment_amount)
            worksheet.write(row, 2, emi.advance_payment)
            worksheet.write(row, 3, emi.remaining_debt)
            worksheet.write(row, 4, emi.payment_status)
            worksheet.write(row, 5, emi.loan_id.loan_bank.name if emi.loan_id.loan_bank else 'No Bank')  # Bank Name
            worksheet.write(row, 6, emi.loan_id.interest_rate)  # Interest Rate
            worksheet.write(row, 7,
                            dict(emi.loan_id._fields['loan_type'].selection).get(emi.loan_id.loan_type))  # Loan Type

            print(f'Emi paid {emi.payment_amount}')
            row += 1

        # Close the workbook and save it to the BytesIO buffer
        workbook.close()

        # Save the Excel file as base64 to store it in Odoo
        excel_file = base64.b64encode(output.getvalue())

        # Create an attachment to store the Excel file
        attachment = self.env['ir.attachment'].create({
            'name': f'EMI_History_{self.loan_no}.xlsx',
            'type': 'binary',
            'datas': excel_file,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'res_model': 'debt.details',
            'res_id': self.id,
        })

        # Return the attachment so the user can download it
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }

    ######### Constrains ##########
    @api.constrains('starting_date', 'first_emi')
    def _check_loan_dates(self):
        for record in self:
            # Ensure that first_emi is not before starting_date
            if record.first_emi and record.first_emi < record.starting_date:
                raise models.ValidationError("The first EMI date cannot be before the loan starting date.")

    @api.constrains('loan_type', 'loan_tenor')
    def _check_tenure_range(self):
        for record in self:
            if record.loan_type == 'personal':
                if not (12 <= record.loan_tenor <= 84):
                    raise models.ValidationError("For Personal Loan, tenure must be between 12 to 84 months.")
            elif record.loan_type == 'debt_consolidation':
                if not (12 <= record.loan_tenor <= 84):
                    raise models.ValidationError("For Debt Consolidation Loan, tenure must be between 12 to 84 months.")
            elif record.loan_type == 'mortgage':
                if not (120 <= record.loan_tenor <= 360):  # 10 to 30 years in months
                    raise models.ValidationError("For Mortgage, tenure must be between 120 to 360 months.")
            elif record.loan_type == 'home_equity':
                if not (60 <= record.loan_tenor <= 360):  # 5 to 30 years in months
                    raise models.ValidationError("For Home Equity Loan, tenure must be between 60 to 360 months.")
            elif record.loan_type == 'student':
                if not (120 <= record.loan_tenor <= 180):  # 10 to 15 years in months
                    raise models.ValidationError("For Student Loan, tenure must be between 120 to 180 months.")
            elif record.loan_type == 'auto':
                if not (12 <= record.loan_tenor <= 84):
                    raise models.ValidationError("For Auto Loan, tenure must be between 12 to 84 months.")
            elif record.loan_type == 'small_business':
                if not (12 <= record.loan_tenor <= 300):
                    raise models.ValidationError("For Small Business Loan, tenure must be between 12 to 300 months.")
            elif record.loan_type == 'credit_builder':
                if record.loan_tenor != 24:
                    raise models.ValidationError("For Credit Builder Loan, tenure must be exactly 24 months.")
            elif record.loan_type == 'payday':
                if not (record.loan_tenor <= 2):  # 2 to 4 weeks, converted to months
                    raise models.ValidationError(
                        "For Payday Loan, tenure must be between 2 to 4 weeks (approximately 0.5 to 1 month).")

    ######### Constrains ##########
    @api.constrains('sanctioned_amount', 'actual_amount')
    def _check_amount(self):
        for record in self:
            # Ensure that first_emi is not before starting_date
            if record.actual_amount and record.sanctioned_amount < record.actual_amount:
                raise models.ValidationError("The Actual Amount cannot be more than Sanctioned Amount.")
