import { CommonModule } from '@angular/common'
import { NgModule } from '@angular/core'
import { InputClearComponent } from '../components/input-clear/input-clear.component'
import { InputNumberComponent } from '../components/input-number/input-number.component'
import { ReactiveFormsModule } from '@angular/forms'
import { NZ_ICONS, NzIconModule } from 'ng-zorro-antd/icon'
import { UserOutline, LockOutline } from '@ant-design/icons-angular/icons'
import { FormsModule } from '@angular/forms'
import { NzInputModule } from 'ng-zorro-antd/input'
import { NzDatePickerModule } from 'ng-zorro-antd/date-picker'
import { NzSelectModule } from 'ng-zorro-antd/select'
import { NzAutocompleteModule } from 'ng-zorro-antd/auto-complete'
import { NzRadioModule } from 'ng-zorro-antd/radio'
import { NzButtonModule } from 'ng-zorro-antd/button'
import { NzAvatarModule } from 'ng-zorro-antd/avatar'
import { NzCommentModule } from 'ng-zorro-antd/comment'
import { NzTableModule } from 'ng-zorro-antd/table'
import { NzTimePickerModule } from 'ng-zorro-antd/time-picker'
import { NzUploadModule } from 'ng-zorro-antd/upload'
import { NzSpinModule } from 'ng-zorro-antd/spin'
import { NzImageModule } from 'ng-zorro-antd/image'
import { NzTypographyModule } from 'ng-zorro-antd/typography'
import { NzTransferModule } from 'ng-zorro-antd/transfer'
import { NzCheckboxModule } from 'ng-zorro-antd/checkbox'
import { NzTabsModule } from 'ng-zorro-antd/tabs'
import { NzModalModule } from 'ng-zorro-antd/modal'
import { NzAlertModule } from 'ng-zorro-antd/alert'
import { NzFormModule } from 'ng-zorro-antd/form'
import { NzPageHeaderModule } from 'ng-zorro-antd/page-header'
import { NzSpaceModule } from 'ng-zorro-antd/space'
import { NzDescriptionsModule } from 'ng-zorro-antd/descriptions'
import { NzDividerModule } from 'ng-zorro-antd/divider'
import { NzPaginationModule } from 'ng-zorro-antd/pagination'
import { NzTagModule } from 'ng-zorro-antd/tag'
import { NzBreadCrumbModule } from 'ng-zorro-antd/breadcrumb'
import { NzEmptyModule } from 'ng-zorro-antd/empty'
import { NzDrawerModule } from 'ng-zorro-antd/drawer'
import { NzInputNumberModule } from 'ng-zorro-antd/input-number'
import { NzPopconfirmModule } from 'ng-zorro-antd/popconfirm'
import { NzTreeModule } from 'ng-zorro-antd/tree'
import { NzTreeViewModule } from 'ng-zorro-antd/tree-view'
import { NzListModule } from 'ng-zorro-antd/list'
import { NzPopoverModule } from 'ng-zorro-antd/popover'
import { NzCollapseModule } from 'ng-zorro-antd/collapse'
import { NzTimelineModule } from 'ng-zorro-antd/timeline'
import { CommentComponent } from '../components/comment/comment.component'
import { PermissionDirective } from '../directives/permission.directive'
import { NzDropDownModule } from 'ng-zorro-antd/dropdown'
import { NzCardModule } from 'ng-zorro-antd/card'

@NgModule({
  providers: [{ provide: NZ_ICONS, useValue: [UserOutline, LockOutline] }],
  imports: [
    InputClearComponent,
    InputNumberComponent,
    CommentComponent,
    PermissionDirective,
  ],
  declarations: [],
  exports: [
    PermissionDirective,
    CommonModule,
    NzCheckboxModule,
    NzInputModule,
    NzIconModule,
    NzAutocompleteModule,
    NzSelectModule,
    NzDatePickerModule,
    NzRadioModule,
    NzButtonModule,
    NzAvatarModule,
    NzCommentModule,
    NzTableModule,
    NzTimePickerModule,
    NzImageModule,
    NzTransferModule,
    NzSpinModule,
    NzTypographyModule,
    NzTabsModule,
    NzModalModule,
    NzFormModule,
    ReactiveFormsModule,
    NzAlertModule,
    NzPageHeaderModule,
    NzDescriptionsModule,
    NzDividerModule,
    NzSpaceModule,
    NzPaginationModule,
    NzTagModule,
    NzBreadCrumbModule,
    FormsModule,
    NzUploadModule,
    InputClearComponent,
    InputNumberComponent,
    NzEmptyModule,
    NzDrawerModule,
    NzPopconfirmModule,
    NzTreeModule,
    NzTreeViewModule,
    NzListModule,
    NzCollapseModule,
    CommentComponent,
    NzTimelineModule,
    NzPopoverModule,
    NzInputNumberModule,
    NzDropDownModule,
    NzCardModule,
  ],
})
export class ShareModule {}
