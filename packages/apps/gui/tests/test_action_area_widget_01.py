"""Focused tests split from test_action_area_widget.py."""

from __future__ import annotations

from ._action_area_widget_support import (
    ActionArea,
    ActionAreaViewModel,
    QApplication,
    QFrame,
    QLabel,
    QPushButton,
    Qt,
    Sizing,
    _combo_data,
    _install_optimization_lookup,
    build_action_area_stylesheet,
    numbering_schemes,
    pytest,
)

pytestmark = pytest.mark.gui
from ._action_area_widget_support import (
    vm as vm,
)
from ._action_area_widget_support import (
    widget as widget,
)


def test_numbering_scheme_projection_filters_locale_and_honors_wildcard() -> None:
    config = {
        "settings": {"order": ["universal", "english", "chinese", "unrestricted"]},
        "schemes": {
            "universal": {"name": "Universal", "locales": ["*"]},
            "english": {"name": "English", "locales": ["en_US"]},
            "chinese": {"name": "Chinese", "locales": ["zh_CN"]},
            "unrestricted": {"name": "Unrestricted"},
        },
    }

    assert numbering_schemes.get_numbering_scheme_items(locale="zh_CN", config_data=config) == [
        ("Universal", "universal"),
        ("Chinese", "chinese"),
        ("Unrestricted", "unrestricted"),
    ]
    assert numbering_schemes.get_numbering_scheme_items(locale="en_US", config_data=config) == [
        ("Universal", "universal"),
        ("English", "english"),
        ("Unrestricted", "unrestricted"),
    ]


class TestConstruction:
    def test_widget_created(self, widget: ActionArea) -> None:
        assert widget is not None

    def test_object_name(self, widget: ActionArea) -> None:
        assert widget.objectName() == "actionAreaRoot"

    def test_initially_hidden(self, widget: ActionArea) -> None:
        assert widget.isVisible() is False

    def test_focus_policy(self, widget: ActionArea) -> None:
        assert widget.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_view_model_access(self, widget: ActionArea) -> None:
        assert widget.view_model is not None
        assert isinstance(widget.view_model, ActionAreaViewModel)

    def test_button_stack_exists(self, widget: ActionArea) -> None:
        assert widget.button_stack is not None
        assert widget.button_stack.count() == 2

    def test_active_actions_render_inside_a_titled_surface(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")

        card = widget.findChild(QFrame, "actionContentCard")
        title = widget.findChild(QLabel, "actionPanelTitle")
        assert card is not None
        assert title is not None and title.text().strip()
        assert widget.findChild(QLabel, "actionPanelSubtitle") is None

    def test_option_change_keeps_existing_widget_tree(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        image_checkbox = widget._image_cb
        assert image_checkbox is not None

        image_checkbox.click()

        assert widget._image_cb is image_checkbox
        assert widget._image_cb.isChecked() == vm.extract_image

    def test_same_type_file_change_keeps_widget_identity_and_focus(
        self,
        qapp: QApplication,
        widget: ActionArea,
        vm: ActionAreaViewModel,
    ) -> None:
        vm.setup_for_document_file("/first.docx")
        image_checkbox = widget._image_cb
        assert image_checkbox is not None
        widget.show()
        image_checkbox.setFocus(Qt.FocusReason.OtherFocusReason)
        qapp.processEvents()

        vm.setup_for_document_file("/second.docx")
        qapp.processEvents()

        assert widget._image_cb is image_checkbox
        assert image_checkbox.hasFocus()
        assert vm.file_path == "/second.docx"

    @pytest.mark.parametrize("locale", ["de_DE", "fr_FR", "ru_RU"])
    def test_long_locale_primary_buttons_use_natural_width(self, qapp: QApplication, locale: str) -> None:
        from docwen_gui.i18n import get_locale, set_locale

        previous_locale = get_locale()
        localized_vm = ActionAreaViewModel()
        localized_widget: ActionArea | None = None
        try:
            set_locale(locale)
            localized_widget = ActionArea(view_model=localized_vm)
            localized_widget.resize(640, 720)
            localized_vm.setup_for_document_file("/localized.docx")
            localized_widget.show()
            qapp.processEvents()

            export_button = localized_widget.document_to_md_button
            assert export_button.minimumWidth() < export_button.maximumWidth()
            assert export_button.width() >= export_button.sizeHint().width()

            localized_vm.setup_for_md_to_document("/localized.md")
            qapp.processEvents()
            generate_button = localized_widget.convert_docx_button
            assert generate_button.minimumWidth() < generate_button.maximumWidth()
            assert generate_button.width() >= generate_button.sizeHint().width()
        finally:
            set_locale(previous_locale)
            if localized_widget is not None:
                localized_widget.close()
                localized_widget.deleteLater()

    def test_stylesheet_overrides_disabled_buttons(self) -> None:
        stylesheet = build_action_area_stylesheet()

        assert "QWidget#actionAreaRoot QPushButton#actionPrimaryButton" in stylesheet
        assert "background-color: palette(highlight)" in stylesheet
        assert "QWidget#actionAreaRoot QPushButton#actionPrimaryButton:disabled" in stylesheet
        assert "background-color: palette(alternate-base)" in stylesheet

    def test_compact_controls_keep_keyboard_and_pointer_targets(
        self,
        widget: ActionArea,
        vm: ActionAreaViewModel,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_document_file("/test.docx")
        controls = [
            widget._cancel_button,
            widget.document_to_md_button,
            widget._image_cb,
            widget._ocr_cb,
            widget._optimize_combo,
            widget.doc_remove_numbering_cb,
            widget.doc_add_numbering_cb,
            widget.doc_numbering_scheme_combo,
        ]

        assert all(control is not None for control in controls)
        for control in controls:
            assert control.minimumHeight() >= Sizing.CONTROL_HEIGHT, (
                f"{control.metaObject().className()} {getattr(control, 'text', lambda: '')()!r} "
                f"has minimumHeight={control.minimumHeight()}"
            )
            assert control.focusPolicy() & Qt.FocusPolicy.TabFocus

    def test_control_metrics_survive_fluent_polish_and_cross_mode_rebuilds(
        self,
        qapp: QApplication,
        widget: ActionArea,
        vm: ActionAreaViewModel,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        setup_steps = (
            lambda: vm.setup_for_document_file("/first.docx"),
            lambda: vm.setup_for_md_to_document("/first.md"),
            lambda: vm.setup_for_md_to_spreadsheet("/first.md"),
            lambda: vm.setup_for_image_file("/first.png"),
            lambda: vm.setup_for_document_file("/second.docx"),
        )
        widget.show()

        for setup in setup_steps:
            setup()
            widget.ensurePolished()
            qapp.processEvents()

            controls = widget._interactive_controls()  # pyright: ignore[reportPrivateUsage]
            assert controls
            undersized = [
                (
                    control.metaObject().className(),
                    getattr(control, "text", lambda: "")(),
                    control.minimumHeight(),
                )
                for control in controls
                if control.minimumHeight() < Sizing.CONTROL_HEIGHT
            ]
            assert undersized == []


class TestVisibility:
    def test_show_panel(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        widget.show_panel()
        assert vm.visible is True

    def test_hide_panel(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        widget.show_panel()
        widget.hide_panel()
        assert vm.visible is False

    def test_show_cancel(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        widget.show_cancel()
        assert vm.cancel_visible is True

    def test_hide_cancel(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        widget.show_cancel()
        widget.hide_cancel()
        assert vm.cancel_visible is False


class TestDocumentToMd:
    @pytest.fixture(autouse=True)
    def setup_vm(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")

    def test_button_assigned(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert widget.document_to_md_button is not None
        assert isinstance(widget.document_to_md_button, QPushButton)

    def test_image_cb_exists(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert widget._image_cb is not None

    def test_ocr_cb_exists(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert widget._ocr_cb is not None

    def test_numbering_cb_exists(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert widget.doc_remove_numbering_cb is not None
        assert widget.doc_add_numbering_cb is not None

    def test_convert_click_emits(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        widget.document_to_md_button.click()
        assert len(emitted) == 1
        assert emitted[0][0] == "md"

    def test_numbering_scheme_combo_uses_all_registered_ids(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        combo = widget.doc_numbering_scheme_combo
        assert combo is not None
        values = [combo.itemData(i) for i in range(combo.count())]
        assert values == [
            scheme_id
            for _label, scheme_id in numbering_schemes.get_numbering_scheme_items(
                config_data=vm.numbering_scheme_config(),
            )
        ]

    def test_numbering_scheme_combo_shows_display_names(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(
            "docwen_gui.widgets.action_area.numbering_schemes.get_numbering_scheme_items",
            lambda locale=None, *, config_data: [
                ("公文标准", "gongwen_standard"),
                ("自定义方案", "custom_scheme"),
            ],
        )
        vm.setup_for_document_file("/localized.docx")
        combo = widget.doc_numbering_scheme_combo
        assert combo is not None
        assert [combo.itemText(i) for i in range(combo.count())] == ["公文标准", "自定义方案"]
        assert [combo.itemData(i) for i in range(combo.count())] == ["gongwen_standard", "custom_scheme"]

    def test_numbering_scheme_combo_reflects_runtime_registry(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(
            "docwen_gui.widgets.action_area.numbering_schemes.get_numbering_scheme_items",
            lambda locale=None, *, config_data: [("Custom One", "custom_one"), ("Custom Two", "custom_two")],
        )
        vm.setup_for_document_file("/registry.docx")
        combo = widget.doc_numbering_scheme_combo
        assert combo is not None
        assert [combo.itemData(i) for i in range(combo.count())] == ["custom_one", "custom_two"]

    def test_optimize_combo_uses_document_scope(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_document_file("/optimized.docx")
        combo = widget._optimize_combo
        assert combo is not None
        assert _combo_data(combo) == [None, "gongwen"]

    def test_optimization_discovery_failure_is_visible_and_does_not_block_conversion(
        self,
        widget: ActionArea,
        vm: ActionAreaViewModel,
    ) -> None:
        from docwen_gui.i18n import t

        vm._main_vm = None  # pyright: ignore[reportPrivateUsage]
        vm.setup_for_document_file("/runtime-unavailable.docx")

        notice = widget.findChild(QLabel, "actionOptimizationUnavailable")
        assert notice is not None
        assert notice.text() == t(
            "action_area.optimization_unavailable",
            "Optimization options are unavailable; standard conversion remains available.",
        )
        assert notice.property("errorCode") == "capability_unavailable"
        assert widget._optimize_combo is None
        assert widget.document_to_md_button.isEnabled()


class TestImageToMd:
    def test_ocr_default_enabled(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_image_file("/test.png")
        assert vm.extract_ocr is True
        assert widget._ocr_cb is not None

    def test_optimize_combo_uses_image_scope(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_image_file("/test.png")
        combo = widget._optimize_combo
        assert combo is not None
        assert _combo_data(combo) == [None, "invoice_cn"]

    def test_primary_boolean_options_share_one_compact_row(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_image_file("/test.png")
        assert widget._image_cb is not None
        assert widget._ocr_cb is not None
        assert widget._image_cb.parentWidget() is widget._ocr_cb.parentWidget()

    def test_selecting_image_invoice_sets_action_and_stays_selected(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_image_file("/test.png")
        combo = widget._optimize_combo
        assert combo is not None
        combo.setCurrentIndex(combo.findData("invoice_cn"))
        assert vm.action_name == "invoice_cn"
        assert widget._optimize_combo is not None
        assert widget._optimize_combo.itemData(widget._optimize_combo.currentIndex()) == "invoice_cn"


class TestLayoutToMd:
    def test_optimize_combo_uses_layout_scope(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_layout_file("/test.pdf")
        combo = widget._optimize_combo
        assert combo is not None
        assert _combo_data(combo) == [None, "invoice_cn"]

    def test_selecting_layout_invoice_sets_action(
        self, widget: ActionArea, vm: ActionAreaViewModel, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_layout_file("/test.pdf")
        combo = widget._optimize_combo
        assert combo is not None
        combo.setCurrentIndex(combo.findData("invoice_cn"))
        assert vm.action_name == "invoice_cn"


class TestOtherToMd:
    def test_no_numbering(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.show_numbering is False

    def test_ready_empty_catalog_hides_optimize(
        self,
        widget: ActionArea,
        vm: ActionAreaViewModel,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        _install_optimization_lookup(monkeypatch)
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.show_optimize is False
        assert widget.findChild(QLabel, "actionOptimizationUnavailable") is None

    def test_button_assigned(self, widget: ActionArea, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert widget.document_to_md_button is not None
