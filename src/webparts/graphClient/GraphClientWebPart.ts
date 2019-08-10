import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphClientWebPart.module.scss';
import * as strings from 'GraphClientWebPartStrings';

// Microsoft Graphへの問い合わせ実行のために追加 パッケージ追加は不要
import { MSGraphClient } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client';

// Microsoft Graphとのやり取りに使う型があるほうがコーディングが楽なので追加
// 要パッケージ追加
// npm install @microsoft/microsoft-graph-types --save-dev
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphClientWebPartProps {
  description: string;
}

export default class GraphClientWebPart extends BaseClientSideWebPart<IGraphClientWebPartProps> {

  public render(): void {

    // 予定一覧を取得
    this.getEvents((error: GraphError, datas: any) => {

      if(error){
        this.domElement.innerHTML = `
          <div class="${ styles.graphClient }">
            <div class="${ styles.container }">
              <div class="${ styles.error }">${ (error)? JSON.stringify(error) : '' }</div>
            </div>
          </div>
        `;
      }
      else {
        let events : MicrosoftGraph.Event[] = datas.value;
        this.domElement.innerHTML = `
          <div class="${ styles.graphClient }">
            <div class="${ styles.container }">
              <div>予定表一覧</div>
              ${
                (events && events.length > 0)?
                  `<table class="${ styles.events }" >
                    <thead>
                      <tr>
                        <th>件名</th>
                        <th>場所</th>
                        <th>日時</th>
                        <th>終日</th>
                        <th>繰り返し</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${events.map((event) => {
                        return `
                          <tr>
                            <td>${ event.subject }</td>
                            <td>${ event.locations.map((location) => {
                              return `
                                <dvi>${ location.displayName }</div>
                              `;
                            }) }</td>
                            <td>${ event.start.dateTime } ～ ${ event.end.dateTime }</td>
                            <td>${ event.isAllDay }</td>
                            <td>${
                              (event.recurrence)? 
                                `<button onclick='javascript:alert(&#x27;${ JSON.stringify(event.recurrence) }&#x27;);' >する</button>` : 
                                'しない' 
                            }</td>
                          </tr>
                        `;
                      })}
                    </tbody>
                  </table>` :
                  '予定がありません'
              }
            </div>
          </div>
        `;
      }
    });
  }

  /**
   * Microsoft Graphからのデータ取得
   * 
   * Microsoft Graphへの問い合わせにはアクセス許可が必要です。
   * そのためconfigフォルダ > package-solution.jsonファイルのsolutionプロパティ内に以下を追記してあります。
   *  "webApiPermissionRequests": [
        {
          "resource": "Microsoft Graph",
          "scope": "Calendars.Read"
        }
      ]
     また、当パッケージをSharePointのアプリカタログサイトに展開した後、
     SharePoint管理センター > APIの管理 画面で
     当パッケージが要求するアクセス許可(Calendars.Read)の承認が必要です。
    
     取得する日付範囲は固定値です。実際のソリューションではシステム日付をフォーマットするなどしてください。
   */
  protected getEvents(callBack : (error: GraphError, datas: any) => void): Promise<any> {
    return this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient) => {
      // ユーザー自身(me)の予定表(calender)から予定(events)を取得
      //  URLパラメータ
      //    $filter=Start/DateTime ge '2019-08-10T00:00:00Z'
      //      予定の開始時刻が指定時刻以上であるという条件のフィルタ。
      //      リクエストヘッダ：Preferで指定のタイムゾーンに従った値でフィルタされる。
      //            
      //  リクエストヘッダ
      //    Prefer: outlook.timezone="Tokyo Standard Time"
      //      "Tokyo Standard Time"はタイムゾーンを表す文字列。指定しないとUTCになる。
      //      以下エンドポイントでサポートされている文字列を検索できる。
      //      https://graph.microsoft.com/v1.0/me/outlook/supportedTimeZones
      return client
        .api("me/calendar/events?$filter=Start/DateTime ge '2019-08-10T00:00:00Z'")
        .headers({
          'Prefer' : 'outlook.timezone="Tokyo Standard Time"'
        })
        .get((error: GraphError, datas: any, rawResponse?: any) => {
          callBack(error, datas);
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
