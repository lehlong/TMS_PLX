import {Component, Input, ViewChild, ElementRef} from '@angular/core';
import {formatDistance} from 'date-fns';
import {CommentService} from '../../../services/components/comment.service';
import {ModuleAttachmentService} from '../../../services/module-attachment.service';
import {CommentFilter} from '../../../models/components/comment.model';
import {PaginationResult} from '../../../models/base.model';
import {TYPE_COMMENT, MODULE_TYPE_COMMENT} from '../../../shared/constants';
import {NzCommentModule} from 'ng-zorro-antd/comment';
import {NzIconModule} from 'ng-zorro-antd/icon';
import {NzAvatarModule} from 'ng-zorro-antd/avatar';
import {NzSpinModule} from 'ng-zorro-antd/spin';
import {NzInputModule} from 'ng-zorro-antd/input';
import {NzImageModule} from 'ng-zorro-antd/image';
import {NzPopconfirmModule} from 'ng-zorro-antd/popconfirm';
import {FormsModule} from '@angular/forms';
import {CommonModule} from '@angular/common';

@Component({
  selector: 'app-comment',
  templateUrl: './comment.component.html',
  styleUrls: ['./comment.component.scss'],
  standalone: true,
  imports: [
    NzIconModule,
    NzCommentModule,
    NzAvatarModule,
    NzSpinModule,
    NzInputModule,
    NzImageModule,
    FormsModule,
    CommonModule,
    NzPopconfirmModule,
  ],
})
export class CommentComponent {
  @ViewChild('scrollContent') scrollContent!: ElementRef;
  @Input() refId: string = '';
  @Input() title: string = 'Bình luận';
  @Input() uploadImage: boolean = true;
  disableText: boolean = false;
  imageUrl: string = '';
  imageUrlSub: string = '';
  loadImage: boolean = false;
  openPrevew: boolean = false;
  attachmentId: string | null = null;
  attachmentIdSub: string | null = null;
  messageChat: string = '';
  messageChatSub: string = '';
  filter = new CommentFilter();
  paginationResult!: PaginationResult;
  loadingContent: boolean = false;
  formatDistance = formatDistance;
  TYPE_COMMENT = TYPE_COMMENT;
  MODULE_TYPE_COMMENT = MODULE_TYPE_COMMENT;

  constructor(private _service: CommentService, private moduleAttachmentService: ModuleAttachmentService) {}

  ngOnInit(): void {}

  getTime(time: string) {
    return formatDistance(new Date(time), new Date());
  }

  ngOnChanges() {
    if (this.refId) {
      this.filter.ReferenceId = this.refId;
      this.loadInit();
    } else {
      this.paginationResult = new PaginationResult();
    }
  }

  onScrollContent() {
    const element = this.scrollContent.nativeElement;
    const atBottom = element.scrollHeight - element.scrollTop === element.clientHeight;

    if (atBottom && this.paginationResult.totalRecord > this.filter.pageSize) {
      this.filter.pageSize = this.filter.pageSize + 10;
      this.loadingContent = true;
      this.loadInit();
    }
  }

  deleteComment(comment: any) {
    this._service.delete(comment.id).subscribe(
      () => {
        if (comment.pId) {
          this.messageChatSub = '';
          this.imageUrlSub = '';
          this.attachmentIdSub = null;
        } else {
          this.messageChat = '';
          this.imageUrl = '';
          this.disableText = false;
          this.loadImage = false;
          this.openPrevew = false;
          this.attachmentId = null;
        }
        this.loadInit();
      },
      (error) => {
        console.log('error: ', error);
      },
    );
  }

  loadInit() {
    this.getAllByReference();
  }

  replyTo(comment: any) {
    this.attachmentIdSub = null;
    this.messageChatSub = '';
    this.imageUrlSub = '';
    this.paginationResult = {
      ...this.paginationResult,
      data: this.paginationResult.data.map((element: any) => {
        return {
          ...element,
          openReply: comment.id === element.id,
          disableText: false,
          openPrevew: false,
          loadImage: false,
        };
      }),
    };

    setTimeout(() => {
      const inputReply = document.getElementById('input-reply') as HTMLInputElement;
      inputReply.focus();
    }, 50);
  }

  getAllByReference() {
    this._service.getAllByReference(this.filter).subscribe(
      (response) => {
        this.paginationResult = {
          ...response,
          data: response?.data.map((item: any) => {
            return {
              ...item,
              openReply: false,
              openPrevew: false,
              loadImage: false,
              disableText: false,
            };
          }),
        };
        this.loadingContent = false;
      },
      (error) => {
        console.log('error: ', error);
      },
    );
  }

  handleImageUpload(event: Event, id: string | null = null, main: boolean = true): void {
    if (main) {
      this.loadImage = true;
      this.openPrevew = true;
      this.messageChat = '';
      this.disableText = true;
    } else {
      this.messageChatSub = '';
      this.paginationResult = {
        ...this.paginationResult,
        data: this.paginationResult.data.map((element: any) => {
          return {
            ...element,
            openPrevew: id === element.id,
            loadImage: id === element.id,
            disableText: true,
          };
        }),
      };
    }

    const inputElement = event.target as HTMLInputElement;
    if (inputElement.files && inputElement?.files.length > 0) {
      const selectedImage = inputElement?.files[0];
      const formData = new FormData();
      formData.append('file', selectedImage, selectedImage.name);
      this.moduleAttachmentService
        .Upload(selectedImage, {
          moduleType: MODULE_TYPE_COMMENT,
          referenceId: this.refId,
        })
        .subscribe(
          ({data}) => {
            if (main) {
              this.imageUrl = data?.url;
              this.attachmentId = data?.id;
              this.loadImage = false;
            } else {
              this.imageUrlSub = data?.url;
              this.attachmentIdSub = data?.id;
              this.paginationResult = {
                ...this.paginationResult,
                data: this.paginationResult.data.map((element: any) => {
                  return {
                    ...element,
                    loadImage: false,
                  };
                }),
              };
            }
            inputElement.value = '';
          },
          (error) => {
            console.log('error: ', error);
          },
        );
    }
  }

  sendMessage(event: any = null, pID: string | null = null) {
    if (!event || event?.key === 'Enter') {
      const attachmentId = pID ? this.attachmentIdSub : this.attachmentId;
      if (!attachmentId) {
        if ((pID && this.messageChatSub == '') || (!pID && this.messageChat == '')) {
          return;
        }
      }
      this._service
        .create({
          pId: pID,
          type: attachmentId ? TYPE_COMMENT.IMAGE : TYPE_COMMENT.TEXT,
          content: pID ? this.messageChatSub : this.messageChat,
          attachmentId: attachmentId,
          referenceId: this.refId,
        })
        .subscribe(
          ({data}) => {
            if (pID) {
              this.attachmentIdSub = null;
              this.messageChatSub = '';
            } else {
              this.disableText = false;
              this.imageUrl = '';
              this.openPrevew = false;
              this.loadImage = false;
              this.attachmentId = null;
              this.messageChat = '';
            }
            this.filter.pageSize = this.filter.pageSize + 1;
            this.loadInit();
          },
          (error) => {
            console.log('error: ', error);
          },
        );
    }
  }

  cleareFile(main: boolean = true) {
    if (main) {
      this.disableText = false;
      this.imageUrl = '';
      this.openPrevew = false;
      this.loadImage = false;
      this.attachmentId = null;
    } else {
      this.attachmentIdSub = null;
      this.paginationResult = {
        ...this.paginationResult,
        data: this.paginationResult.data.map((element: any) => {
          return {
            ...element,
            imageUrlSub: '',
            openPrevew: false,
            loadImage: false,
            disableText: false,
          };
        }),
      };
    }
  }
}
