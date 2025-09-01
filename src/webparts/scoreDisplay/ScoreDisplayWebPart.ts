import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './ScoreDisplayWebPart.module.scss';
import * as strings from 'ScoreDisplayWebPartStrings';

export interface IScoreDisplayWebPartProps {
  description: string;
}

interface Student {
  studentId: string;
  name: string;
  chinese: number;
  english: number;
  math: number;
  science: number;
  total: number;
  average: number;
}

export default class ScoreDisplayWebPart extends BaseClientSideWebPart<IScoreDisplayWebPartProps> {

  private students: Student[] = [];
  private filteredStudents: Student[] = [];

  public render(): void {
    try {
      this.domElement.innerHTML = `
      <div class="${styles.scoreDisplay || 'scoreDisplay'}">
        <div class="${styles.header || 'header'}">
          <h1>學生成績管理系統</h1>
          <p>使用 SharePoint Framework 開發的學生成績顯示系統</p>
        </div>
        
        <div class="${styles.content || 'content'}">
          <div id="loading" class="${styles.loading || 'loading'}">
            <div class="${styles.spinner || 'spinner'}"></div>
            <p>正在載入學生成績資料...</p>
          </div>
          
          <div id="error" class="${styles.error || 'error'}" style="display: none;">
            <p id="errorMessage">載入失敗，請稍後再試</p>
          </div>
          
          <div id="mainContent" style="display: none;">
            <!-- 篩選器 -->
            <div class="${styles.filters || 'filters'}">
              <input 
                id="searchInput"
                type="text" 
                placeholder="搜尋學生姓名或學號..." 
                class="${styles.filterInput || 'filterInput'}"
              >
              <select id="subjectFilter" class="${styles.filterInput || 'filterInput'}">
                <option value="">所有科目</option>
                <option value="chinese">中文</option>
                <option value="english">英文</option>
                <option value="math">數學</option>
                <option value="science">科學</option>
              </select>
              <button id="exportBtn" class="${styles.exportBtn || 'exportBtn'}">匯出 CSV</button>
            </div>
            
            <!-- 成績表格 -->
            <div class="${styles.tableContainer || 'tableContainer'}">
              <table id="gradesTable" class="${styles.gradesTable || 'gradesTable'}">
                <thead>
                  <tr>
                    <th>學號</th>
                    <th>姓名</th>
                    <th>中文</th>
                    <th>英文</th>
                    <th>數學</th>
                    <th>科學</th>
                    <th>總分</th>
                    <th>平均分</th>
                  </tr>
                </thead>
                <tbody id="tableBody">
                </tbody>
              </table>
            </div>
            
            <!-- 統計資訊 -->
            <div class="${styles.summaryStats || 'summaryStats'}">
              <div class="${styles.statCard || 'statCard'}">
                <h3>總學生數</h3>
                <p id="totalStudents" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>平均分數</h3>
                <p id="averageScore" class="${styles.statNumber || 'statNumber'}">0.00</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>最高分數</h3>
                <p id="highestScore" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
              <div class="${styles.statCard || 'statCard'}">
                <h3>最低分數</h3>
                <p id="lowestScore" class="${styles.statNumber || 'statNumber'}">0</p>
              </div>
            </div>
          </div>
        </div>
      </div>`;

      this._loadStudentData();
      this._setupEventListeners();
    } catch (error) {
      console.error('Error in render method:', error);
      this.domElement.innerHTML = `
        <div style="padding: 20px; text-align: center; color: #d13438;">
          <h2>載入失敗</h2>
          <p>學生成績系統載入時發生錯誤，請稍後再試。</p>
          <p style="font-size: 12px; color: #666;">錯誤詳情: ${error.message}</p>
        </div>
      `;
    }
  }

  private _loadStudentData(): void {
    try {
      // 模擬Excel數據
      this.students = [
        { studentId: 'S001', name: '張小明', chinese: 85, english: 92, math: 88, science: 90, total: 355, average: 88.75 },
        { studentId: 'S002', name: '李小華', chinese: 78, english: 85, math: 92, science: 87, total: 342, average: 85.5 },
        { studentId: 'S003', name: '王小美', chinese: 92, english: 88, math: 85, science: 94, total: 359, average: 89.75 },
        { studentId: 'S004', name: '陳小強', chinese: 80, english: 90, math: 95, science: 82, total: 347, average: 86.75 },
        { studentId: 'S005', name: '林小芳', chinese: 88, english: 85, math: 90, science: 88, total: 351, average: 87.75 },
        { studentId: 'S006', name: '黃小偉', chinese: 75, english: 92, math: 87, science: 90, total: 344, average: 86.0 },
        { studentId: 'S007', name: '劉小玲', chinese: 90, english: 88, math: 93, science: 85, total: 356, average: 89.0 },
        { studentId: 'S008', name: '吳小傑', chinese: 82, english: 85, math: 88, science: 92, total: 347, average: 86.75 },
        { studentId: 'S009', name: '趙小雅', chinese: 95, english: 90, math: 89, science: 91, total: 365, average: 91.25 },
        { studentId: 'S010', name: '孫小龍', chinese: 87, english: 93, math: 91, science: 89, total: 360, average: 90.0 }
      ];

      // 模擬載入延遲
      setTimeout(() => {
        try {
          this.filteredStudents = [...this.students];
          this._hideLoading();
          this._renderTable();
          this._updateStats();
        } catch (error) {
          console.error('Error in data loading timeout:', error);
          this._showError('數據載入失敗');
        }
      }, 1000);
    } catch (error) {
      console.error('Error in _loadStudentData:', error);
      this._showError('初始化數據失敗');
    }
  }

  private _setupEventListeners(): void {
    try {
      const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      const subjectFilter = this.domElement.querySelector('#subjectFilter') as HTMLSelectElement;
      const exportBtn = this.domElement.querySelector('#exportBtn') as HTMLButtonElement;

      if (searchInput) {
        searchInput.addEventListener('input', () => this._filterStudents());
      }
      if (subjectFilter) {
        subjectFilter.addEventListener('change', () => this._filterStudents());
      }
      if (exportBtn) {
        exportBtn.addEventListener('click', () => this._exportToCSV());
      }
    } catch (error) {
      console.error('Error setting up event listeners:', error);
    }
  }

  private _filterStudents(): void {
    try {
      const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
      const subjectFilter = this.domElement.querySelector('#subjectFilter') as HTMLSelectElement;
      
      const searchTerm = searchInput?.value.toLowerCase() || '';
      const selectedSubject = subjectFilter?.value || '';

      this.filteredStudents = this.students.filter(student => {
        const matchesSearch = student.name.toLowerCase().indexOf(searchTerm) !== -1 || 
                             student.studentId.toLowerCase().indexOf(searchTerm) !== -1;
        
        let matchesSubject = true;
        if (selectedSubject) {
          const subjectScore = student[selectedSubject as keyof Student] as number;
          matchesSubject = subjectScore >= 80;
        }

        return matchesSearch && matchesSubject;
      });

      this._renderTable();
      this._updateStats();
    } catch (error) {
      console.error('Error in _filterStudents:', error);
    }
  }

  private _renderTable(): void {
    try {
      const tableBody = this.domElement.querySelector('#tableBody');
      if (!tableBody) return;

      tableBody.innerHTML = this.filteredStudents.map(student => `
        <tr>
          <td>${this._escapeHtml(student.studentId)}</td>
          <td>${this._escapeHtml(student.name)}</td>
          <td>${student.chinese}</td>
          <td>${student.english}</td>
          <td>${student.math}</td>
          <td>${student.science}</td>
          <td>${student.total}</td>
          <td>${student.average.toFixed(2)}</td>
        </tr>
      `).join('');
    } catch (error) {
      console.error('Error in _renderTable:', error);
    }
  }

  private _updateStats(): void {
    try {
      const totalStudents = this.domElement.querySelector('#totalStudents');
      const averageScore = this.domElement.querySelector('#averageScore');
      const highestScore = this.domElement.querySelector('#highestScore');
      const lowestScore = this.domElement.querySelector('#lowestScore');

      if (this.filteredStudents.length === 0) {
        if (totalStudents) totalStudents.textContent = '0';
        if (averageScore) averageScore.textContent = '0.00';
        if (highestScore) highestScore.textContent = '0';
        if (lowestScore) lowestScore.textContent = '0';
        return;
      }

      const avg = this.filteredStudents.reduce((sum, student) => sum + student.average, 0) / this.filteredStudents.length;
      const max = Math.max(...this.filteredStudents.map(student => student.average));
      const min = Math.min(...this.filteredStudents.map(student => student.average));

      if (totalStudents) totalStudents.textContent = this.filteredStudents.length.toString();
      if (averageScore) averageScore.textContent = avg.toFixed(2);
      if (highestScore) highestScore.textContent = max.toString();
      if (lowestScore) lowestScore.textContent = min.toString();
    } catch (error) {
      console.error('Error in _updateStats:', error);
    }
  }

  private _exportToCSV(): void {
    try {
      const headers = ['學號', '姓名', '中文', '英文', '數學', '科學', '總分', '平均分'];
      const csvContent = [
        headers.join(','),
        ...this.filteredStudents.map(student => 
          [student.studentId, student.name, student.chinese, student.english, student.math, student.science, student.total, student.average.toFixed(2)].join(',')
        )
      ].join('\n');

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', 'student_grades.csv');
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      console.error('Error in _exportToCSV:', error);
      alert('匯出失敗，請稍後再試');
    }
  }

  private _hideLoading(): void {
    try {
      const loading = this.domElement.querySelector('#loading') as HTMLElement;
      const mainContent = this.domElement.querySelector('#mainContent') as HTMLElement;
      
      if (loading) loading.style.display = 'none';
      if (mainContent) mainContent.style.display = 'block';
    } catch (error) {
      console.error('Error in _hideLoading:', error);
    }
  }

  private _showError(message: string): void {
    try {
      const errorDiv = this.domElement.querySelector('#error') as HTMLElement;
      const errorMessage = this.domElement.querySelector('#errorMessage') as HTMLElement;
      const loading = this.domElement.querySelector('#loading') as HTMLElement;
      
      if (loading) loading.style.display = 'none';
      if (errorDiv) errorDiv.style.display = 'block';
      if (errorMessage) errorMessage.textContent = message;
    } catch (error) {
      console.error('Error in _showError:', error);
    }
  }

  private _escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
