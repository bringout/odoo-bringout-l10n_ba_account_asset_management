# Copyright 2009-2018 Noviat
# Copyright 2019 Tecnativa - Pedro M. Baeza
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import calendar
import logging
from datetime import date
from functools import reduce
from sys import exc_info
from traceback import format_exception

from dateutil.relativedelta import relativedelta

from odoo import _, api, fields, models
from odoo.exceptions import UserError
from odoo.osv import expression

_logger = logging.getLogger(__name__)


class AccountAsset(models.Model):
    _inherit = "account.asset"
   
    status_bosanski = fields.Char(
        string="Status po bosanski",
        compute="_compute_status_bosanski",
        readonly=False,
        store=False,
    )

    def _compute_status_bosanski(self):
        def _bosanski(term):
            match term:
                case "draft":
                    return "pripr."
                case "open":
                    return "aktiv."
                case "close":
                    return "amort."
                case "removed":
                    return "otpis."
                case _:
                    return term
        
        for rec in self:
            rec.status_bosanski = _bosanski(rec.state)


    @api.model
    def _xls_acquisition_fields(self):
        """
        Update list in custom module to add/drop columns or change order
        """
        return [
            "account",
            "name",
            "code",
            "date_start",
            "purchase_value",
        ]

    @api.model
    def _xls_active_fields(self):
        """
        Update list in custom module to add/drop columns or change order
        """
        return [
            "account",
            "name",
            "code",
            "date_start",
            "purchase_value",
            "period_start_value",
            "period_depr",
            "period_end_depr",
            "period_end_value", # sadašnja vrijednost (vrijednost na kraju perioda)
            "method",
            "method_number",
            "prorata",
            "state",
        ]

    @api.model
    def _xls_removal_fields(self):
        """
        Update list in custom module to add/drop columns or change order
        """
        return [
            "account",
            "name",
            "code",
            "date_remove",
            "purchase_value",
        ]

    #@api.model
    #def _xls_asset_template(self):
    #    """
    #    Template updates
    #
    #    """
    #    return {}
    #
    #@api.model
    #def _xls_acquisition_template(self):
    #    """
    #    Template updates
    #
    #    """
    #    return {}
    #
    #@api.model
    #def _xls_active_template(self):
    #    """
    #    Template updates
    #
    #    """
    #    return {}
    #
    #@api.model
    #def _xls_removal_template(self):
    #    """
    #    Template updates
    #
    #    """
    #    return {}
