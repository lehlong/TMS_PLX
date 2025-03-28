import { ShareModule } from './../../shared/share-module/index'
import { Component, OnInit, ViewChild } from '@angular/core'
import {
  NzFormatEmitEvent,
  NzTreeComponent,
  NzTreeNodeOptions,
} from 'ng-zorro-antd/tree'
import { OrganizeService } from '../../services/system-manager/organize.service'
import { FormGroup, NonNullableFormBuilder, Validators } from '@angular/forms'
import { GlobalService } from '../../services/global.service'
import { NzMessageService } from 'ng-zorro-antd/message'
import { ADMIN_RIGHTS } from '../../shared/constants'

@Component({
  selector: 'app-organize',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './organize.component.html',
  styleUrl: './organize.component.scss',
})
export class OrganizeComponent implements OnInit {
  @ViewChild('treeCom', { static: false }) treeCom!: NzTreeComponent
  searchValue = ''
  nodes: any = []
  originalNodes: any[] = []
  visible: boolean = false
  edit: boolean = false
  nodeCurrent!: any
  titleParent: string = ''

  validateForm: FormGroup = this.fb.group({
    id: ['', [Validators.required]],
    name: ['', [Validators.required]],
    pId: ['', [Validators.required]],
    children: [null],
    orderNumber: [null],
    isActive: [true, [Validators.required]],
  })
  ADMIN_RIGHTS = ADMIN_RIGHTS
  constructor(
    private _service: OrganizeService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Cấu trúc tổ chức',
        path: 'system-manager/organize',
      },
    ])
  }

  ngOnInit(): void {
    this.getOrg()
  }

  getOrg() {
    this._service.GetOrgTree().subscribe((res) => {
      this.nodes = [res]
      this.originalNodes = [res]
      console.log("node: ", this.nodes);

    })
  }

  nzEvent(event: NzFormatEmitEvent): void {
    // console.log(event);
  }

  onDrop(event: any) { }

  onDragStart(event: any) { }

  onClick(node: any) {
    this.edit = true
    this.visible = true
    this.nodeCurrent = node?.origin
    this.titleParent = node.parentNode?.origin?.title || ''
    this.validateForm.setValue({
      id: this.nodeCurrent?.id,
      name: this.nodeCurrent?.name,
      pId: this.nodeCurrent?.pId,
      children: [],
      orderNumber: this.nodeCurrent?.orderNumber,
      isActive: this.nodeCurrent?.isActive,
    })
  }

  close() {
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.searchValue = ''
    this.getOrg()
    this.nodes = [...this.originalNodes]
  }

  resetForm() {
    this.validateForm.reset()
  }

  openCreateChild(node: any) {
    this.close()
    this.edit = false
    this.visible = true
    this.validateForm.get('pId')?.setValue(node?.origin.id)
    this.validateForm.get('orderNumber')?.setValue(null)
    this.validateForm.get('children')?.setValue([])
    this.validateForm.get('isActive')?.setValue(true)
  }

  openCreate() {
    this.close()
    this.edit = false
    this.visible = true
    this.validateForm.get('pId')?.setValue(this.nodeCurrent?.id || 'R')
    this.validateForm.get('children')?.setValue([])
    this.validateForm.get('orderNumber')?.setValue(null)
    this.validateForm.get('isActive')?.setValue(true)
  }
  isIdExist(id: string, node: any): boolean {
    if (node.id === id) {
      return true
    }
    if (node.children) {
      for (const child of node.children) {
        if (this.isIdExist(id, child)) {
          return true
        }
      }
    }
    return false
  }
  submitForm() {
    if (!this.validateForm.valid) {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty()
          control.updateValueAndValidity({ onlySelf: true })
        }
      })
      return
    }
    console.log('Form Data:', this.validateForm.getRawValue())
    if (this.edit) {
      this._service.Update(this.validateForm.getRawValue()).subscribe({
        next: (data) => {
          this.getOrg()
        },
        error: (response) => {
          console.log(response)
        },
      })
    } else {
      console.log('data', this.validateForm.getRawValue())
      const formData = this.validateForm.getRawValue()
      const newId = formData.id
      const idExists = this.nodes.some((node: any) =>
        this.isIdExist(newId, node),
      )
      if (idExists) {
        this.message.error(
          `Mã đơn vị ${newId} đã được sử dụng, vui lòng nhập lại`,
        )
        return
      }
      this._service.Insert(formData).subscribe({
        next: (data) => {
          this.getOrg()
        },
        error: (response) => {
          console.log(response)
        },
      })
    }
  }

  updateOrderTree() {
    const treeData = this.treeCom
      .getTreeNodes()
      .map((node) => this.mapNode(node))
    console.log('data', treeData[0])
    this._service.UpdateOrderTree(treeData[0]).subscribe({
      next: (data) => {
        this.getOrg()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  private mapNode(node: any): any {
    const children = node.children
      ? node.children.map((child: any) => this.mapNode(child))
      : []
    return {
      id: node.origin.id,
      pId: node.parentNode?.key || '-',
      name: node.origin.name,
      children: children,
      isActive: node.origin.isActive,
    }
  }

  deleteItem(node: any) {
    if (node.children && node.children.length > 0) {
      // Thông báo rằng không thể xóa vì node có children
      this.message.error(
        'Không được phép xóa Cấu trúc tổ chức Cha khi còn các thành phần con',
      )
      return // Dừng quá trình xóa
    }
    this._service.Delete(node.origin.id).subscribe({
      next: (data) => {
        this.getOrg()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  searchTables(searchValue: string) {
    const filterNode = (node: NzTreeNodeOptions): boolean => {
      if (node.title.toLowerCase().includes(searchValue.toLowerCase())) {
        return true
      } else if (node.children) {
        node.children = node.children.filter((child) => filterNode(child))
        return node.children.length > 0
      }
      return false
    }

    if (!searchValue) {
      this.nodes = [...this.originalNodes]
    } else {
      this.nodes = this.originalNodes
        .map((node) => ({ ...node }))
        .filter((node) => filterNode(node))
    }
  }
}
