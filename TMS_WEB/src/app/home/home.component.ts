import { Component, OnInit } from '@angular/core'
import { ShareModule } from '../shared/share-module';
declare var google: any
@Component({
  selector: 'app-home',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './home.component.html',
  styleUrl: './home.component.scss',
})
export class HomeComponent implements OnInit {
  ngOnInit(): void {
    //this.dashBoard1();
    //this.dashBoard2();
  }

  dashBoard1(){
    google.charts.load('current', {'packages':['corechart']});
    google.charts.setOnLoadCallback(drawVisualization);

    function drawVisualization() {
      // Some raw data (not necessarily accurate)
      var data = google.visualization.arrayToDataTable([
        ['Month', 'Bolivia', 'Ecuador', 'Madagascar', 'Papua New Guinea', 'Rwanda', 'Average'],
        ['2004/05',  165,      938,         522,             998,           450,      614.6],
        ['2005/06',  135,      1120,        599,             1268,          288,      682],
        ['2006/07',  157,      1167,        587,             807,           397,      623],
        ['2007/08',  139,      1110,        615,             968,           215,      609.4],
        ['2008/09',  136,      691,         629,             1026,          366,      569.6]
      ]);

      var options = {
        title : 'Monthly Coffee Production by Country',
        vAxis: {title: 'Cups'},
        hAxis: {title: 'Month'},
        seriesType: 'bars',
        series: {5: {type: 'line'}}
      };

      var chart = new google.visualization.ComboChart(document.getElementById('chart_div'));
      chart.draw(data, options);
    }
  }


  dashBoard2(){
    google.charts.load('current', {packages: ['corechart', 'bar']});
google.charts.setOnLoadCallback(drawAnnotations);

function drawAnnotations() {
      var data = new google.visualization.DataTable();
      data.addColumn('timeofday', 'Time of Day');
      data.addColumn('number', 'Motivation Level');
      data.addColumn({type: 'string', role: 'annotation'});
      data.addColumn('number', 'Energy Level');
      data.addColumn({type: 'string', role: 'annotation'});

      data.addRows([
        [{v: [8, 0, 0], f: '8 am'},   1, '1',  .25, '.2'],
        [{v: [9, 0, 0], f: '9 am'},   2, '2',   .5, '.5'],
        [{v: [10, 0, 0], f:'10 am'},  3, '3',    1,  '1'],
        [{v: [11, 0, 0], f: '11 am'}, 4, '4', 2.25,  '2'],
        [{v: [12, 0, 0], f: '12 pm'}, 5, '5', 2.25,  '2'],
        [{v: [13, 0, 0], f: '1 pm'},  6, '6',    3,  '3'],
        [{v: [14, 0, 0], f: '2 pm'},  7, '7', 3.25,  '3'],
        [{v: [15, 0, 0], f: '3 pm'},  8, '8',    5,  '5'],
        [{v: [16, 0, 0], f: '4 pm'},  9, '9',  6.5,  '6'],
        [{v: [17, 0, 0], f: '5 pm'}, 10, '10',  10, '10'],
      ]);

      var options = {
        title: 'Motivation and Energy Level Throughout the Day',
        annotations: {
          alwaysOutside: true,
          textStyle: {
            fontSize: 14,
            color: '#000',
            auraColor: 'none'
          }
        },
        hAxis: {
          title: 'Time of Day',
          format: 'h:mm a',
          viewWindow: {
            min: [7, 30, 0],
            max: [17, 30, 0]
          }
        },
        vAxis: {
          title: 'Rating (scale of 1-10)'
        }
      };

      var chart = new google.visualization.ColumnChart(document.getElementById('chart_div_2'));
      chart.draw(data, options);
    }
  }
}
