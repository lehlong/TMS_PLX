import {
  Component,
  EventEmitter,
  Input,
  Output,
  SimpleChanges,
} from '@angular/core'
import { ShareModule } from '../../../shared/share-module'

@Component({
  selector: 'app-menu-right',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './menu-right.component.html',
  styleUrl: './menu-right.component.scss',
})
export class MenuRightComponent {
  @Input() menuDetail!: any
  @Output() childEvent = new EventEmitter<any>()
  nodes!: any
  constructor() { }

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

  sendDataToParent() {
    const listAccountGroupRight = this.getCheckedNodes(this.nodes).map(
      (element: any) => {
        return {
          rightId: element?.key,
        }
      },
    )
    this.childEvent.emit(listAccountGroupRight)
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['menuDetail']) {
      console.log('Received menuDetail:', changes['menuDetail'].currentValue);
      this.nodes = this.mapTreeNodes(
        changes['menuDetail'].currentValue.treeRight
      );
      console.log('Mapped nodes:', this.nodes);
      this.sendDataToParent();
    }
  }

  mapTreeNodes(data: any): any[] {
    console.log('Processing node:', data);
    return data.children
      ? data.children.map((node: any) => {
        console.log('Mapping child node:', node);
        return {
          title: node.id + '-' + node.name,
          key: node.id,
          checked: node.isChecked,
          expanded: true,
          children: this.mapTreeNodes(node),
        };
      })
      : [];
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
}
