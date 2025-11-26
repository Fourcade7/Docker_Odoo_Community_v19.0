from odoo import api, fields, models


class EmiPayment(models.Model):
    _name = 'debt.emi.history'
    _description = 'EMI Payment History'

    loan_id = fields.Many2one('debt.details', string='Loan', required=True, ondelete='cascade')
    due_date = fields.Date(string='Due Date', required=True)
    payment_amount = fields.Float(string='Payment Amount', required=True)
    advance_payment = fields.Float(string='Advance Payment', default=0.0)  # New field for advance payment
    remaining_debt = fields.Float(string='Remaining Debt', compute='_compute_remaining_debt', store=True)
    payment_status = fields.Selection([
        ('paid', 'Paid'),
        ('missed', 'Missed')
    ], string='Payment Status', default='paid')

    @api.depends('payment_amount', 'advance_payment', 'loan_id.total_debt', 'due_date')
    def _compute_remaining_debt(self):
        for record in self:
            # Get all EMI records for this loan ordered by due date
            emi_records = self.env['debt.emi.history'].search([
                ('loan_id', '=', record.loan_id.id),
                ('due_date', '<=', record.due_date)
            ], order='due_date')

            # Calculate the total amount paid up to this EMI record (including advance payments)
            total_paid = 0.0
            advance_applied = False  # Flag to track if an advance payment has been applied

            for emi in emi_records:
                total_paid += emi.payment_amount  # Only include EMI payments

                # If there's an advance payment, apply it directly to remaining debt
                if emi.advance_payment :
                    total_paid += emi.advance_payment  # Add the advance payment only once


            # Calculate the remaining debt after this payment
            remaining_debt = record.loan_id.total_debt - total_paid

            # If remaining debt is less than 0, set it to 0 (avoid negative debt)
            if remaining_debt < 0:
                remaining_debt = 0

            # Ensure the remaining debt is 0 if the advance payment fully matches it
            if record.advance_payment >= remaining_debt:
                remaining_debt = 0  # Set remaining debt to 0 if advance payment fully covers the debt

            record.remaining_debt = remaining_debt
            print(f"Remaining debt for EMI on {record.due_date}: {record.remaining_debt}")
