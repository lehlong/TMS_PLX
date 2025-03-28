import { ShareModule } from './../../shared/share-module/index'
import { ChangeDetectorRef, Component, ViewChild } from '@angular/core'
import {
  NzFormatEmitEvent,
  NzTreeComponent,
  NzTreeNodeOptions,
} from 'ng-zorro-antd/tree'
import { FormGroup, NonNullableFormBuilder, Validators } from '@angular/forms'
import { GlobalService } from '../../services/global.service'
import { MenuService } from '../../services/system-manager/menu.service'
import { MenuRightComponent } from './menu-right/menu-right.component'
import { NzMessageService } from 'ng-zorro-antd/message'
import { ADMIN_RIGHTS } from '../../shared/constants'

@Component({
  selector: 'app-menu',
  standalone: true,
  imports: [ShareModule, MenuRightComponent],
  templateUrl: './menu.component.html',
  styleUrl: './menu.component.scss',
})
export class MenuComponent {
  @ViewChild('treeCom', { static: false }) treeCom!: NzTreeComponent
  searchValue = ''
  nodes: any = []
  originalNodes: any[] = []
  visible: boolean = false
  edit: boolean = false
  nodeCurrent!: any
  titleParent: string = ''
  tabIndex: number = 0
  menuRight: any

  validateForm: FormGroup = this.fb.group({
    id: ['', [Validators.required]],
    name: ['', [Validators.required]],
    pId: ['', [Validators.required]],
    url: [''],
    orderNumber: [null],
    rightId: [''],
    icon: [''],
    children: [null],
  })
  ADMIN_RIGHTS = ADMIN_RIGHTS
  constructor(
    private _service: MenuService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
    private cdr: ChangeDetectorRef,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách menu',
        path: 'system-manager/menu',
      },
    ])
  }

  ngOnInit(): void {
    this.getMenus()
  }

  getMenus() {
    this._service.GetMenuTree().subscribe((res) => {
      this.nodes = [res]
      this.originalNodes = [res]
    })
  }

  nzEvent(event: NzFormatEmitEvent): void {
    // console.log(event);
  }

  onDrop(event: any) { }

  onDragStart(event: any) { }

  onClick(node: any) {
    this._service
      .GetMenuWithTreeRight({
        menuId: node?.origin?.id,
      })
      .subscribe((res) => {
        this.nodeCurrent = res
        this.edit = true
        this.visible = true
        this.titleParent = node.parentNode?.origin?.title || ''
        this.validateForm.setValue({
          id: this.nodeCurrent?.id,
          name: this.nodeCurrent?.name,
          pId: this.nodeCurrent?.pId,
          url: this.nodeCurrent?.url,
          icon: this.nodeCurrent?.icon,
          orderNumber: this.nodeCurrent?.orderNumber,
          rightId: this.nodeCurrent?.rightId || '',
          children: [],
        })
      })
  }

  close() {
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.searchValue = ''
    this.getMenus()
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
  }

  openCreate() {
    this.close()
    this.edit = false
    this.visible = true
    this.validateForm.get('pId')?.setValue(this.nodeCurrent?.id || 'MNU')
    this.validateForm.get('orderNumber')?.setValue(null)
    this.validateForm.get('children')?.setValue([])
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
    if (this.tabIndex == 0) {
      if (this.edit) {
        this._service.Update(this.validateForm.getRawValue()).subscribe({
          next: (data) => {
            this.getMenus()
          },
          error: (response) => {
            console.log(response)
          },
        })
      } else {
        const formData = this.validateForm.getRawValue()
        const newId = formData.id
        const idExists = this.nodes.some((node: any) =>
          this.isIdExist(newId, node),
        )
        if (idExists) {
          this.message.error(
            `Mã menu ${newId} đã được sử dụng, vui lòng nhập lại`,
          )
          return
        }
        this._service.Insert(this.validateForm.getRawValue()).subscribe({
          next: (data) => {
            this.getMenus()
          },
          error: (response) => {
            console.log(response)
          },
        })
      }
    } else {
      const param = {
        id: this.nodeCurrent.id,
        name: this.nodeCurrent.name,
        pId: this.nodeCurrent.pId,
        url: this.nodeCurrent.url,
        orderNumber: this.nodeCurrent.orderNumber,
        RightReferences: this.menuRight,
        icon: this.nodeCurrent.icon,
        children: this.nodeCurrent.children,
      }
      this._service.Update(param).subscribe({
        next: (data) => {
          this.getMenus()
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
    this._service.UpdateOrderTree(treeData[0]).subscribe({
      next: (data) => {
        this.getMenus()
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
      pId: node.parentNode?.key || 'MNU',
      rightId: node.origin?.rightId,
      name: node.origin.name,
      url: node.origin.url || '',
      icon: node.origin.icon || '',
      children: children,
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
        this.getMenus()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  onSelectedTab(event: any) {
    this.tabIndex = event
  }

  handleChildEvent(data: any) {
    this.menuRight = data
  }
  searchTables(searchValue: string) {
    const filterNode = (node: NzTreeNodeOptions): NzTreeNodeOptions | null => {
      const isMatch = node.title
        .toLowerCase()
        .includes(searchValue.toLowerCase())

      if (node.children) {
        const filteredChildren = node.children
          .map((child) => filterNode(child))
          .filter((child) => child !== null) as NzTreeNodeOptions[]

        if (isMatch || filteredChildren.length > 0) {
          return {
            ...node,
            children: filteredChildren,
          }
        }
      } else if (isMatch) {
        return node
      }

      return null
    }

    if (!searchValue) {
      this.nodes = [...this.originalNodes]
    } else {
      this.nodes = this.originalNodes
        .map((node) => filterNode(node))
        .filter((node) => node !== null) as NzTreeNodeOptions[]
    }

    // Force view update
    this.cdr.detectChanges()
  }
}
