import {Injectable} from '@angular/core';
import * as signalR from '@aspnet/signalr';
import {BehaviorSubject} from 'rxjs';
import {CommonService} from '../common.service';
import {environment} from '../../../environments/environment';
import {Observable} from 'rxjs';
@Injectable({
  providedIn: 'root',
})
export class DeviceConnectionService {
  URL_SIGNALR = `${environment.baseApiUrl.replace('/api', '')}/SystemTrace`;
  hubDeviceConnectionMessage = 'SYSTEM_TRACE';
  hubDeviceConnection: BehaviorSubject<string>;
  hubUrl: string = '';
  connection: any = null;
  constructor(private _commonService: CommonService) {
    {
      this.hubUrl = this.URL_SIGNALR;
      this.hubDeviceConnection = new BehaviorSubject<string>('');
    }
  }

  getAllDevice(): Observable<any> {
    return this._commonService.get(`SystemTrace/GetAll`, {});
  }
  Insert(params: any): Observable<any> {
    return this._commonService.post(`SystemTrace/Insert`, params);
  }

  Update(params: any): Observable<any> {
    return this._commonService.put(`SystemTrace/Update`, params);
  }

  delete(code: string | number): Observable<any> {
    return this._commonService.delete(`SystemTrace/Delete/${code}`);
  }

  public async initiateSignalrConnection(): Promise<void> {
    try {
      if (this.connection) {
        await this.connection.stop();
      }
      const token = localStorage.getItem('token') || 'null';
      const options: signalR.IHttpConnectionOptions = {
        accessTokenFactory: () => {
          return token;
        },
        transport: signalR.HttpTransportType.ServerSentEvents,
      };

      this.connection = new signalR.HubConnectionBuilder().withUrl(this.hubUrl, options).build();

      await this.connection.start();
      this.setSignalrClientMethods();
    } catch (error) {
      console.log('error: ', error);
      if (error == 'Error: Unauthorized') {
      } else {
        console.log(`SignalR connection error: ${error}`);
      }
    }
  }

  private setSignalrClientMethods(): void {
    this.connection.on(this.hubDeviceConnectionMessage, (message: any) => {
      this.hubDeviceConnection.next(message);
    });
  }
}
