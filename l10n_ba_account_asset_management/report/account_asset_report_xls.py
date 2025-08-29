# Copyright 2025-2025 bring.out doo Sarajevo
# Copyright 2009-2019 Noviat
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import logging

from odoo import _, models
from odoo.exceptions import UserError

from odoo.addons.report_xlsx_helper.report.report_xlsx_format import (
    FORMATS,
    XLS_HEADERS,
)

_logger = logging.getLogger(__name__)


IR_TRANSLATION_NAME = "account.asset.report"


class AssetReportXlsx(models.AbstractModel):
    _inherit = "report.account_asset_management.asset_report_xls"
    _description = "Dynamic XLS asset report generator Bosna i Hercegovina"


    
    def _get_asset_template(self):

        def _bosanski(term):
            match term:
                case "open":
                    return "aktivno"
                case "close":
                    return "amortizovano"
                case "removed":
                    return "otpisano"
                case _:
                    return term
         
        asset_template = {
            "account": {
                "header": {"type": "string", "value": _("Konto")},
                "asset": {
                    "type": "string",
                    "value": self._render(
                        "asset.profile_id.account_asset_id.code or ''"
                    ),
                },
                "totals": {"type": "string", "value": _("Ukupno")},
                "width": 10,
            },
            "name": {
                "header": {"type": "string", "value": _("Ime")},
                "asset_group": {
                    "type": "string",
                    "value": self._render("group.name or ''"),
                },
                "asset": {"type": "string", "value": self._render("asset.name")},
                "width": 50,
            },
            "code": {
                "header": {"type": "string", "value": _("Referenca")},
                "asset_group": {
                    "type": "string",
                    "value": self._render("group.code or ''"),
                },
                "asset": {"type": "string", "value": self._render("asset.code or ''")},
                "width": 20,
            },
            "date_start": {
                "header": {"type": "string", "value": _("Aktivirano")},
                "asset": {
                    "value": self._render("asset.date_start or ''"),
                    "format": FORMATS["format_tcell_date_left"],
                },
                "width": 12,
            },
            "date_remove": {
                "header": {"type": "string", "value": _("Datum otpisa")},
                "asset": {
                    "value": self._render("asset.date_remove or ''"),
                    "format": FORMATS["format_tcell_date_left"],
                },
                "width": 20,
            },
            "depreciation_base": {
                "header": {
                    "type": "string",
                    "value": _("Osnova"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "number",
                    "value": self._render('0'),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render("0"),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "number",
                    "value": self._render("0"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "salvage_value": {
                "header": {
                    "type": "string",
                    "value": _("Rezervna vrijednost"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "number",
                    "value": self._render('0'),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render("0"),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "number",
                    "value": self._render("0"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "purchase_value": {
                "header": {
                    "type": "string",
                    "value": _("Nabavna vrijednost"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "number",
                    "value": self._render('group_entry["_purchase_value"]'),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render("asset.purchase_value"),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "formula",
                    "value": self._render("purchase_total_formula"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "period_start_value": {
                "header": {
                    "type": "string",
                    "value": _("Osnovica"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "number",
                    "value": self._render('group_entry["_period_start_value"]'),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render('asset_entry["_period_start_value"]'),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "formula",
                    "value": self._render("period_start_total_formula"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "period_depr": {
                "header": {
                    "type": "string",
                    "value": _("Amort. period"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "formula",
                    "value": self._render("period_diff_formula"),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "formula",
                    "value": self._render("period_diff_formula"),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "formula",
                    "value": self._render("period_diff_formula"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "period_end_value": {
                "header": {
                    "type": "string",
                    "value": _("Sadašnja vrijednost"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "number",
                    "value": self._render('group_entry["_period_end_value"]'),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render('asset_entry["_period_end_value"]'),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "formula",
                    "value": self._render("period_end_total_formula"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "period_end_depr": {
                "header": {
                    "type": "string",
                    "value": _("Amort. ukupno"),
                    "format": FORMATS["format_theader_yellow_right"],
                },
                "asset_group": {
                    "type": "formula",
                    "value": self._render("total_depr_formula"),
                    "format": FORMATS["format_theader_blue_amount_right"],
                },
                "asset": {
                    "type": "formula",
                    "value": self._render("total_depr_formula"),
                    "format": FORMATS["format_tcell_amount_right"],
                },
                "totals": {
                    "type": "formula",
                    "value": self._render("total_depr_formula"),
                    "format": FORMATS["format_theader_yellow_amount_right"],
                },
                "width": 18,
            },
            "method": {
                "header": {
                    "type": "string",
                    "value": _("Metoda obračuna"),
                    "format": FORMATS["format_theader_yellow_center"],
                },
                "asset": {
                    "type": "string",
                    "value": self._render("asset.method or ''"),
                    "format": FORMATS["format_tcell_center"],
                },
                "width": 20,
            },
            "method_number": {
                "header": {
                    "type": "string",
                    "value": _("Broj godina/mjeseci"),
                    "format": FORMATS["format_theader_yellow_center"],
                },
                "asset": {
                    "type": "number",
                    "value": self._render("asset.method_number"),
                    "format": FORMATS["format_tcell_integer_center"],
                },
                "width": 20,
            },
            "prorata": {
                "header": {
                    "type": "string",
                    "value": _("Proporcionalno"),
                    "format": FORMATS["format_theader_yellow_center"],
                },
                "asset": {
                    "type": "boolean",
                    "value": self._render("asset.prorata"),
                    "format": FORMATS["format_tcell_center"],
                },
                "width": 20,
            },
            "state": {
                "header": {
                    "type": "string",
                    "value": _("Status"),
                    "format": FORMATS["format_theader_yellow_center"],
                },
                "asset": {
                    "type": "string",
                    "value": self._render("asset.status_bosanski"),
                    "format": FORMATS["format_tcell_center"],
                },
                "width": 8,
            },
        }

        asset_template.update(self.env["account.asset"]._xls_asset_template())

        return asset_template



    def _get_title(self, wiz, report, frmt="normal"):

        prefix = "{} - {}".format(wiz.date_from, wiz.date_to)
        if report == "acquisition":
            if frmt == "normal":
                title = prefix + " : " + _("Novonabavljena")
            else:
                title = "NOVA"
        elif report == "active":
            if frmt == "normal":
                title = prefix + " : " + _("Aktivna")
            else:
                title = "AKTIVNA"
        else:
            if frmt == "normal":
                title = prefix + " : " + _("Otpisana")
            else:
                title = "OTPISANA"
        return title

    def _report_title(self, ws, row_pos, ws_params, data, wiz):
        return self._write_ws_title(ws, row_pos, ws_params)

    def _empty_report(self, ws, row_pos, ws_params, data, wiz):
        report = ws_params["report_type"]
        if report == "acquisition":
            suffix = _("Novonabavljena")
        elif report == "active":
            suffix = _("Aktivna")
        else:
            suffix = _("Uklonjena")
        no_entries = _("No") + " " + suffix
        ws.write_string(row_pos, 0, no_entries, FORMATS["format_left_bold"])

    def _get_assets(self, wiz, data):
        """Add the selected assets, both grouped and ungrouped, to `data`"""
        dom = [
            ("date_start", "<=", wiz.date_to),
            "|",
            ("date_remove", "=", False),
            ("date_remove", ">=", wiz.date_from),
        ]

        parent_group = wiz.asset_group_id
        if parent_group:

            def _child_get(parent):
                groups = [parent]
                children = parent.child_ids
                children = children.sorted(lambda r: r.code or r.name)
                for child in children:
                    if child in groups:
                        raise UserError(
                            _(
                                "Nekonzistentna struktura izvještaja."
                                "\nMolimo ispravite grupu sredstva '{group}' (id {id})"
                            ).format(group=child.name, id=child.id)
                        )
                    groups.extend(_child_get(child))
                return groups

            groups = _child_get(parent_group)
            dom.append(("group_ids", "in", [x.id for x in groups]))

        if not wiz.draft:
            dom.append(("state", "!=", "draft"))
        assets = self.env["account.asset"].search(dom)
        grouped_assets = {}
        self._group_assets(assets, parent_group, grouped_assets)
        data.update(
            {
                "assets": assets,
                "grouped_assets": grouped_assets,
            }
        )

    @staticmethod
    def acquisition_filter(wiz, asset):
        return asset.date_start >= wiz.date_from

    @staticmethod
    def active_filter(wiz, asset):
        return True

    @staticmethod
    def removal_filter(wiz, asset):
        return (
            asset.date_remove
            and asset.date_remove >= wiz.date_from
            and asset.date_remove <= wiz.date_to
        )

    def _group_assets(self, assets, group, grouped_assets):
        if group:
            group_assets = assets.filtered(lambda r: group in r.group_ids)
        else:
            group_assets = assets
        group_assets = group_assets.sorted(
            lambda r: (r.date_start or "", r.code or "", r.name)
        )
        grouped_assets[group] = {"assets": group_assets}
        for child in group.child_ids:
            self._group_assets(assets, child, grouped_assets[group])

    def _create_report_entries(
        self, ws_params, wiz, entries, group, group_val, error_dict
    ):
        report = ws_params["report_type"]

        def asset_filter(asset):
            filt = getattr(self, "{}_filter".format(report))
            return filt(wiz, asset)

        def _has_assets(group, group_val):
            assets = group_val.get("assets")
            assets = assets.filtered(asset_filter)
            if assets:
                return True
            for child in group.child_ids:
                if _has_assets(child, group_val[child]):
                    return True
            return False

        assets = group_val.get("assets")
        assets = assets.filtered(asset_filter)

        # remove empty entries
        if not _has_assets(group, group_val):
            return

        asset_entries = []
        group_entry = {
            "_purchase_value": 0.0,
            "_period_start_value": 0.0,
            "_period_end_value": 0.0,
            "group": group,
        }
        for asset in assets:
            asset_entry = {"asset": asset}
            group_entry["_purchase_value"] += asset.purchase_value
            dls_all = asset.depreciation_line_ids.filtered(
                lambda r: r.type == "depreciate"
            )
            dls_all = dls_all.sorted(key=lambda r: r.line_date)
            if not dls_all and asset.method_number:
                error_dict["no_table"] += asset
            # period_start_value
            dls = dls_all.filtered(lambda r: r.line_date <= wiz.date_from)
            if dls:
                value_depreciated = dls[-1].depreciated_value + dls[-1].amount
            else:
                value_depreciated = 0.0
            asset_entry["_period_start_value"] = (
                asset.purchase_value - value_depreciated
            )
            group_entry["_period_start_value"] += asset_entry["_period_start_value"]
            # period_end_value
            dls = dls_all.filtered(lambda r: r.line_date <= wiz.date_to)
            if dls:
                value_depreciated = dls[-1].depreciated_value + dls[-1].amount
            else:
                value_depreciated = 0.0
            asset_entry["_period_end_value"] = (
                asset.purchase_value - value_depreciated
            )
            group_entry["_period_end_value"] += asset_entry["_period_end_value"]

            asset_entries.append(asset_entry)

        todos = []
        for g in group.child_ids:
            if _has_assets(g, group_val[g]):
                todos.append(g)

        entries.append(group_entry)
        entries.extend(asset_entries)
        for todo in todos:
            self._create_report_entries(
                ws_params, wiz, entries, todo, group_val[todo], error_dict
            )

    def _asset_report(self, workbook, ws, ws_params, data, wiz):
        report = ws_params["report_type"]

        ws.set_portrait()
        ws.fit_to_pages(1, 0)
        ws.set_header(XLS_HEADERS["xls_headers"]["standard"])
        ws.set_footer(XLS_HEADERS["xls_footers"]["standard"])

        wl = ws_params["wanted_list"]
        if "account" not in wl:
            raise UserError(
                _(
                    "'Konto' je obavezna stavka "
                    "'_xls_%s_fields' liste !"
                )
                % report
            )

        self._set_column_width(ws, ws_params)

        row_pos = 0
        row_pos = self._report_title(ws, row_pos, ws_params, data, wiz)

        def asset_filter(asset):
            filt = getattr(self, "{}_filter".format(report))
            return filt(wiz, asset)

        assets = data["assets"].filtered(asset_filter)

        if not assets:
            return self._empty_report(ws, row_pos, ws_params, data, wiz)

        row_pos = self._write_line(
            ws,
            row_pos,
            ws_params,
            col_specs_section="header",
            default_format=FORMATS["format_theader_yellow_left"],
        )

        ws.freeze_panes(row_pos, 0)

        row_pos_start = row_pos
        purchase_value_pos = "purchase_value" in wl and wl.index("purchase_value")

        period_start_value_pos = "period_start_value" in wl and wl.index(
            "period_start_value"
        )
        period_end_value_pos = "period_end_value" in wl and wl.index("period_end_value")

        entries = []
        root = wiz.asset_group_id
        root_val = data["grouped_assets"][root]
        error_dict = {
            "no_table": self.env["account.asset"],
            "dups": self.env["account.asset"],
        }

        self._create_report_entries(ws_params, wiz, entries, root, root_val, error_dict)

        # traverse entries in reverse order to calc totals
        for i, entry in enumerate(reversed(entries)):
            if "group" in entry:
                parent = entry["group"].parent_id
                for parent_entry in reversed(entries[: -i - 1]):
                    if "group" in parent_entry and parent_entry["group"] == parent:
                        parent_entry["_purchase_value"] += entry["_purchase_value"]

                        parent_entry["_period_start_value"] += entry[
                            "_period_start_value"
                        ]
                        parent_entry["_period_end_value"] += entry["_period_end_value"]
                        continue

        processed = []
        for entry in entries:

            period_start_value_cell = period_start_value_pos and self._rowcol_to_cell(
                row_pos, period_start_value_pos
            )
            period_end_value_cell = period_end_value_pos and self._rowcol_to_cell(
                row_pos, period_end_value_pos
            )
            period_end_value_cell = period_end_value_pos and self._rowcol_to_cell(
                row_pos, period_end_value_pos
            )
            purchase_value_cell = purchase_value_pos and self._rowcol_to_cell(
                row_pos, purchase_value_pos
            )

            period_diff_formula = period_end_value_cell and (
                period_start_value_cell + "-" + period_end_value_cell
            )

            total_depr_formula = period_end_value_cell and (
                purchase_value_cell + "-" + period_end_value_cell
            )

            if "group" in entry:
                row_pos = self._write_line(
                    ws,
                    row_pos,
                    ws_params,
                    col_specs_section="asset_group",
                    render_space={
                        "group": entry["group"],
                        "group_entry": entry,
                        "period_diff_formula": period_diff_formula,
                        "total_depr_formula": total_depr_formula,
                    },
                    default_format=FORMATS["format_theader_blue_left"],
                )

            else:
                asset = entry["asset"]
                if asset in processed:
                    error_dict["dups"] += asset
                    continue
                else:
                    processed.append(asset)
                row_pos = self._write_line(
                    ws,
                    row_pos,
                    ws_params,
                    col_specs_section="asset",
                    render_space={
                        "asset": entry["asset"],
                        "asset_entry": entry,
                        "period_diff_formula": period_diff_formula,
                        "total_depr_formula": total_depr_formula,
                    },
                    default_format=FORMATS["format_tcell_left"],
                )

        purchase_total_formula = purchase_value_pos and self._rowcol_to_cell(
            row_pos_start, purchase_value_pos
        )
        asset_total_formula = purchase_value_pos and self._rowcol_to_cell(
            row_pos_start, purchase_value_pos
        )
        period_start_total_formula = period_start_value_pos and self._rowcol_to_cell(
            row_pos_start, period_start_value_pos
        )
        period_end_total_formula = period_end_value_pos and self._rowcol_to_cell(
            row_pos_start, period_end_value_pos
        )
        period_start_value_cell = period_start_value_pos and self._rowcol_to_cell(
            row_pos, period_start_value_pos
        )
        period_end_value_cell = period_end_value_pos and self._rowcol_to_cell(
            row_pos, period_end_value_pos
        )

        period_diff_formula = period_end_value_cell and (
            period_start_value_cell + "-" + period_end_value_cell
        )
        total_depr_formula = period_end_value_cell and (
            purchase_total_formula + "-" + period_end_value_cell
        )

        row_pos = self._write_line(
            ws,
            row_pos,
            ws_params,
            col_specs_section="totals",
            render_space={
                "purchase_total_formula": purchase_total_formula,
                "asset_total_formula": asset_total_formula,
                "period_start_total_formula": period_start_total_formula,
                "period_end_total_formula": period_end_total_formula,
                "period_diff_formula": period_diff_formula,
                "total_depr_formula": total_depr_formula,
            },
            default_format=FORMATS["format_theader_yellow_left"],
        )

        for k in error_dict:
            if error_dict[k]:
                if k == "no_table":
                    reason = _("Nedostaje tabela amortizacije")
                elif k == "dups":
                    reason = _("Duple izvještajne")
                else:
                    reason = _("Undetermined error")
                row_pos += 1
                err_msg = _("Sredstva koja se trebaju korigovati") + ": "
                err_msg += "%s" % [x[1] for x in error_dict[k].name_get()]
                err_msg += " - " + _("Razlog") + ": " + reason
                ws.write_string(row_pos, 0, err_msg, FORMATS["format_left_bold"])
