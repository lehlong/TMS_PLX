import { ShareModule } from '../../shared/share-module/index'
import { Component, AfterViewInit, OnInit } from '@angular/core'
import { GlobalService } from '../../services/global.service'
import { ChangeDetectorRef } from '@angular/core'
import { CKEditorModule } from '@ckeditor/ckeditor5-angular'
import {
  ClassicEditor,
  AccessibilityHelp,
  Alignment,
  Autosave,
  Bold,
  Essentials,
  FontBackgroundColor,
  FontColor,
  FontFamily,
  FontSize,
  GeneralHtmlSupport,
  Heading,
  Indent,
  IndentBlock,
  Italic,
  Link,
  Paragraph,
  SelectAll,
  ShowBlocks,
  SourceEditing,
  SpecialCharacters,
  Subscript,
  Superscript,
  Table,
  TableCaption,
  TableCellProperties,
  TableColumnResize,
  TableProperties,
  TableToolbar,
  Underline,
  Undo,
  EditorConfig,
} from 'ckeditor5'
import { ConfigTemplateService } from '../../services/system-manager/config-template.service'
import { FormGroup, NonNullableFormBuilder, Validators } from '@angular/forms'
import { ADMIN_RIGHTS } from '../../shared/constants'

@Component({
  selector: 'config-template-email',
  standalone: true,
  imports: [ShareModule, CKEditorModule],
  templateUrl: './config-template-email.component.html',
  styleUrls: ['./config-template-email.component.scss'],
})
export class ConfixTemplateEmailComponent implements AfterViewInit, OnInit {
  validateForm: FormGroup = this.fb.group({
    code: [''],
    name: ['', [Validators.required]],
    htmlSource: [''],
    type: ['EMAIL'],
    title: [''],
    isActive: [true, [Validators.required]],
  })

  public Editor = ClassicEditor
  public isLayoutReady = false
  public config: EditorConfig = {}
  ADMIN_RIGHTS = ADMIN_RIGHTS
  constructor(
    private fb: NonNullableFormBuilder,
    private _service: ConfigTemplateService,
    private globalService: GlobalService,
    private changeDetector: ChangeDetectorRef,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Cấu hình Template EMAIL',
        path: 'system-manager/confixtemplate-email',
      },
    ])
  }

  title: any = 'Cấu hình template EMAIL'
  edit: boolean = false
  isSubmit: boolean = false
  tabs: any = []
  isCancelModalVisible: boolean = false
  model: any = {
    code: '',
    name: '',
    htmlSource: '',
    type: "EMAIL",
    title: '',
    isActive: true
  }
  indexTab: any = 0
  data: any = []

  ngOnInit(): void {
    this.getAll()
    this.ngAfterViewInit()
    console.log(this.config.initialData);
  }

  getAll() {
    this.isSubmit = false
    this._service.getall().subscribe({
    next: (data) => {
        this.data = data.filter((item: any) => item.type === "EMAIL")
        console.log(this.data);

      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  newTab(): void {
    this.isCancelModalVisible = true
    this.edit = false
  }

  closeTab({ index }: { index: number }): void {
    this.tabs.splice(index, 1);
  }

  handleCancelOk() {
    this.tabs.push(this.model.name);
    console.log(this.tabs);
    this.isCancelModalVisible = false
    this.edit = false
  }

  handleCancelModal() {
    this.isCancelModalVisible = false
    // this.name = ''
    this.edit = true
  }

  trackByItemId(index: number, item: any) {
    console.log(item.code);
  }

  submitForm(): void {
    this.isSubmit = true
    this.isCancelModalVisible = false
    if (this.edit) {
      if (this.validateForm.valid) {
        const formData = this.validateForm.getRawValue()
        console.log(formData);

        this._service.updateConfigTemplate(formData).subscribe({
          next: (data) => {
            this.getAll()
          },
          error: (response) => {
            console.log(response)
          },
        })
      }
    }
    else if (this.model.name) {
      const formData = this.model
      this._service.createConfigTemplate(formData).subscribe({
        next: (data) => {
          this.getAll()
        },
        error: (response) => {
          console.log(response)
        },
      })
    } else {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty()
          control.updateValueAndValidity({ onlySelf: true })
        }
      })
    }
  }


  openEdit(data: any): void {

    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      htmlSource: data.htmlSource,
      title: data.title,
      type: data.type,
      isActive: data.isActive,
    })
    setTimeout(() => {
      this.edit = true
    }, 2000)
  }

  addPram(event: Event, pram: string) {
    event.preventDefault();
    var html = this.validateForm.get('htmlSource')?.value
    var htmlSource = html + ' ' + pram
    this.validateForm.get('htmlSource')?.setValue(htmlSource);
  }

  public ngAfterViewInit(): void {
    this.config = {
      toolbar: {
        items: [
          'undo',
          'redo',
          '|',
          'sourceEditing',
          'showBlocks',
          '|',
          'heading',
          '|',
          'fontSize',
          'fontFamily',
          'fontColor',
          'fontBackgroundColor',
          '|',
          'bold',
          'italic',
          'underline',
          'subscript',
          'superscript',
          '|',
          'specialCharacters',
          'link',
          'insertTable',
          '|',
          'alignment',
          '|',
          'outdent',
          'indent',
        ],
        shouldNotGroupWhenFull: false,
      },
      plugins: [
        AccessibilityHelp,
        Alignment,
        Autosave,
        Bold,
        Essentials,
        FontBackgroundColor,
        FontColor,
        FontFamily,
        FontSize,
        GeneralHtmlSupport,
        Heading,
        Indent,
        IndentBlock,
        Italic,
        Link,
        Paragraph,
        SelectAll,
        ShowBlocks,
        SourceEditing,
        SpecialCharacters,
        Subscript,
        Superscript,
        Table,
        TableCaption,
        TableCellProperties,
        TableColumnResize,
        TableProperties,
        TableToolbar,
        Underline,
        Undo,
      ],
      fontFamily: {
        supportAllValues: true,
      },
      fontSize: {
        options: [10, 12, 14, 'default', 18, 20, 22],
        supportAllValues: true,
      },
      heading: {
        options: [
          {
            model: 'paragraph',
            title: 'Paragraph',
            class: 'ck-heading_paragraph',
          },
          {
            model: 'heading1',
            view: 'h1',
            title: 'Heading 1',
            class: 'ck-heading_heading1',
          },
          {
            model: 'heading2',
            view: 'h2',
            title: 'Heading 2',
            class: 'ck-heading_heading2',
          },
          {
            model: 'heading3',
            view: 'h3',
            title: 'Heading 3',
            class: 'ck-heading_heading3',
          },
          {
            model: 'heading4',
            view: 'h4',
            title: 'Heading 4',
            class: 'ck-heading_heading4',
          },
          {
            model: 'heading5',
            view: 'h5',
            title: 'Heading 5',
            class: 'ck-heading_heading5',
          },
          {
            model: 'heading6',
            view: 'h6',
            title: 'Heading 6',
            class: 'ck-heading_heading6',
          },
        ],
      },
      htmlSupport: {
        allow: [
          {
            name: /^.*$/,
            styles: true,
            attributes: true,
            classes: true,
          },
        ],
      },
      initialData:this.data.htmlSource,
      link: {
        addTargetToExternalLinks: true,
        defaultProtocol: 'https://',
        decorators: {
          toggleDownloadable: {
            mode: 'manual',
            label: 'Downloadable',
            attributes: {
              download: 'file',
            },
          },
        },
      },
      placeholder: 'Type or paste your content here!',
      table: {
        contentToolbar: [
          'tableColumn',
          'tableRow',
          'mergeTableCells',
          'tableProperties',
          'tableCellProperties',
        ],
      },
    }

    this.isLayoutReady = true
    this.changeDetector.detectChanges()
  }




}



