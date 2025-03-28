import { Component, Input, OnInit } from '@angular/core'
import { FormBuilder, FormGroup, Validators } from '@angular/forms'
import { AccountGroupService } from '../../../services/system-manager/account-group.service'
import { ShareModule } from '../../../shared/share-module'
import { DropdownService } from '../../../services/dropdown/dropdown.service'
import { NzFormatEmitEvent, NzTreeNode } from 'ng-zorro-antd/tree'
import { RightService } from '../../../services/system-manager/right.service'
import { AuthService } from '../../../services/auth.service'
import { GlobalService } from '../../../services/global.service'
import { ACCOUNT_GROUP_RIGHTS } from '../../../shared/constants/index'
@Component({
  selector: 'app-account-group-edit',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './account-group-edit.component.html',
  styleUrls: ['./account-group-edit.component.scss'],
})
export class AccountGroupEditComponent implements OnInit {
  @Input() close: () => void = () => { }
  @Input() reset: () => void = () => { }
  @Input() visible: boolean = false
  @Input() id: string | number = 0
  @Input() loading: boolean = false
  nameShow: string = ''
  listAccount: any = []
  searchValue = ''
  nodes: any[] = []
  nodesConstant: any[] = []
  initialCheckedNodes: any[] = []
  validateForm: FormGroup
  ACCOUNT_GROUP_RIGHTS = ACCOUNT_GROUP_RIGHTS
  constructor(
    private accountGroupService: AccountGroupService,
    private rightService: RightService,
    private fb: FormBuilder,
    private dropdownService: DropdownService,
    private authService: AuthService,
    private globalService: GlobalService,
  ) {
    this.validateForm = this.fb.group({
      name: ['', [Validators.required]],
      notes: [''],
      isActive: [true, [Validators.required]],
    })
  }

  ngOnInit(): void {
    if (this.id) {
      this.loadDetail(this.id)
    }
    this.getRight()
  }

  onDrop(event: any): void {
    // Handle drop event
  }

  onDragStart(event: any): void {
    // Handle drag start event
  }

  onClick(event: any): void {
    this.visible = true
    const data = event.node.origin
    this.validateForm.setValue({
      name: data?.name || '',
      notes: data?.notes || '',
      isActive: data?.isActive || true,
    })
  }

  getRight() {
    this.rightService.GetRightTree().subscribe((res) => {
      this.nodes = this.mapTreeNodes(res)
    })
  }
  mapTreeNodes(data: any): any[] {
    return data.children
      ? data.children.map((node: any) => ({
        title: node.name,
        key: node.id,
        checked: node.isChecked,
        expanded: true,
        children: this.mapTreeNodes(node),
      }))
      : []
  }

  onNodeCheckChange(node: any): void {
    node.checked = !node.checked
  }

  getCheckedNodes(nodes: any[]): any[] {
    let checkedNodes: any[] = []
    for (let node of nodes) {
      if (node.checked) {
        checkedNodes.push(node)
      }
      if (node.children) {
        checkedNodes = checkedNodes.concat(this.getCheckedNodes(node.children))
      }
    }
    return checkedNodes
  }
  onCheckBoxChange(event: any): void {
    const checkedNode = event.node
    const nodes = this.flattenKeys(this.nodesConstant)

    if (checkedNode.isChecked) {
      if (nodes.includes(checkedNode.key)) {
        checkedNode.origin.InChecked = false
        checkedNode.origin.OutChecked = false
      } else {
        checkedNode.origin.InChecked = true
      }
      this.checkParents(checkedNode)
      this.checkChildren(checkedNode)
    } else {
      if (nodes.includes(checkedNode.key)) {
        checkedNode.origin.OutChecked = true
        checkedNode.origin.InChecked = false
      } else {
        checkedNode.origin.OutChecked = false
        checkedNode.origin.InChecked = false
      }
      this.uncheckChildren(checkedNode)
    }
  }

  flattenKeys(data: any) {
    return data.reduce((keys: any, item: any) => {
      if (item.checked) {
        keys.push(item.key)
      }
      if (item.children && item.children.length > 0) {
        keys.push(...this.flattenKeys(item.children))
      }
      return keys
    }, [])
  }

  private checkParents(node: NzTreeNode): void {
    const parentNode = node.parentNode
    if (parentNode) {
      parentNode.isChecked = true
      this.checkParents(parentNode)
    }
  }

  private checkChildren(node: NzTreeNode): void {
    if (node.children) {
      node.children.forEach((child) => {
        child.isChecked = true
        this.checkChildren(child)
      })
    }
  }

  private uncheckChildren(node: NzTreeNode): void {
    if (node.children) {
      node.children.forEach((child) => {
        child.isChecked = false
        this.uncheckChildren(child)
      })
    }
  }

  nzEvent(event: NzFormatEmitEvent): void { }

  change(value: any): void { }

  getAllAccount(listUser: any = []): void {
    this.dropdownService.GetAllAccount().subscribe({
      next: (data) => {
        this.listAccount = data.map((account: any) => ({
          ...account,
          direction: listUser.some(
            (user: any) => user?.userName === account?.userName,
          )
            ? 'right'
            : 'left',
          title: `${account?.userName} ${account?.fullName}`,
        }))
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  loadDetail(id: number | string): void {
    this.accountGroupService.GetDetail(id).subscribe({
      next: (data) => {
        this.nameShow = data.name
        this.validateForm.patchValue({
          name: data.name ? data.name : '',
          notes: data.notes ? data.notes : '',
          isActive: data.isActive,
        })
        this.listAccount = data.account_AccountGroups.map(
          (accountGroup: any) => ({
            ...accountGroup.account,
            direction: 'right',
            title: `${accountGroup.account.userName} ${accountGroup.account.fullName}`,
          }),
        )
        this.getAllAccount(data?.account_AccountGroups)
        this.initialCheckedNodes = data.listAccountGroupRight.map(
          (node: any) => node.rightId,
        )
        this.nodes = this.mapTreeNodes(data.treeRight)
        this.nodesConstant = [...this.mapTreeNodes(data.treeRight)]
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  submitForm(): void {
    const account_AccountGroups = this.listAccount?.reduce(
      (result: any, item: any) => {
        if (item?.direction === 'right') {
          return [
            ...result,
            {
              id: item?.id,
              userName: item?.userName,
            },
          ]
        }
        return result
      },
      [],
    )

    const listAccountGroupRight = this.getCheckedNodes(this.nodes).map(
      (element: any) => ({
        rightId: element.key,
      }),
    )

    if (this.validateForm.valid) {
      const updateData = {
        id: this.id,
        ...this.validateForm.value,
      }
      this.accountGroupService
        .Update({ ...updateData, account_AccountGroups, listAccountGroupRight })
        .subscribe({
          next: () => {
            if (this.globalService.getUserInfo().userName) {
              this.authService
                .getRightOfUser({
                  userName: this.globalService.getUserInfo().userName,
                })
                .subscribe({
                  next: (rights) => {
                    this.globalService.setRightData(
                      JSON.stringify(rights || []),
                    )
                  },
                  error: (error) => {
                    console.error('Get right of user failed:', error)
                  },
                })
            }
            this.reset()
          },
          error: (response: any) => {
            console.log('hehe',response)
           
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

  closeDrawer(): void {
    this.close()
    this.resetForm()
  }

  resetForm(): void {
    this.validateForm.reset()
  }
}
