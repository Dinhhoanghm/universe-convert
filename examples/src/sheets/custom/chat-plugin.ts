import { Plugin, Inject, Injector, ICommandService, IUniverInstanceService, UniverInstanceType } from "@univerjs/core";
import { ComponentManager, IUIController } from "@univerjs/ui";
import { SetRangeValuesCommand } from "@univerjs/sheets";
import { html, render } from "lit";

interface ChatMessage {
  text: string;
  type: 'user' | 'bot';
  timestamp: Date;
}

interface SimpleAIConfig {
  aiApiKey: string; // OpenAI compatible key
}

export class ChatPlugin extends Plugin {
  static override pluginName = "ChatPlugin";
  
  private messages: ChatMessage[] = [];
  private chatContainer: HTMLElement | null = null;
  private isTyping: boolean = false;
  private aiConfig: SimpleAIConfig = {
    aiApiKey: ''
  };

  constructor(
    _config: unknown,
    @Inject(Injector) protected readonly _injector: Injector,
    @Inject(ComponentManager) private readonly _componentManager: ComponentManager,
    @Inject(IUIController) private readonly _uiController: IUIController,
    @Inject(ICommandService) private readonly _commandService: ICommandService,
    @Inject(IUniverInstanceService) private readonly _univerInstanceService: IUniverInstanceService
  ) {
    super();
    this.initializeAIConfig();
  }

  private initializeAIConfig() {
    const savedKey = localStorage.getItem('univer_ai_api_key');
    if (savedKey) {
      this.aiConfig.aiApiKey = savedKey;
      setTimeout(() => this.addBotMessage('✅ AI sẵn sàng! Hãy thử: "create revenue table", "sum column B", "put hello in B3"'), 400);
    } else {
      setTimeout(() => this.addBotMessage('🔑 Vui lòng nhập API key trước: gõ "set ai key YOUR_KEY_HERE"'), 400);
    }
  }

  override onReady() {
    this.initializeChatPanel();
  }

  override dispose() {
    if (this.chatContainer) {
      this.chatContainer.remove();
    }
  }

  private initializeChatPanel() {
    this.createFloatingChatPanel();
    this.addToolbarButton();
    (window as any).showUniverChat = () => this.showChat();
  }

  private addToolbarButton() {
    const toolbarButton = document.createElement('button');
    toolbarButton.innerHTML = '💬';
    toolbarButton.title = 'Toggle Chat';
    toolbarButton.style.cssText = `
      background: none; border: none; cursor: pointer; font-size: 18px;
      padding: 8px; margin: 0 4px; border-radius: 4px; transition: background-color 0.2s;
    `;
    
    toolbarButton.addEventListener('mouseover', () => {
      toolbarButton.style.backgroundColor = 'rgba(0,0,0,0.1)';
    });
    
    toolbarButton.addEventListener('mouseout', () => {
      toolbarButton.style.backgroundColor = 'transparent';
    });
    
    toolbarButton.addEventListener('click', () => {
      this.toggleChat();
    });

    setTimeout(() => {
      const toolbar = document.querySelector('.univer-toolbar') || 
                    document.querySelector('[data-range-selector="toolbar"]') ||
                    document.querySelector('.toolbar') ||
                    document.querySelector('.univer-header');
      
      if (toolbar) {
        toolbar.appendChild(toolbarButton);
      } else {
        toolbarButton.style.cssText += `
          position: fixed; top: 10px; right: 10px; z-index: 999;
          background: #f0f0f0; border: 1px solid #ccc; border-radius: 6px;
          padding: 8px 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        `;
        document.body.appendChild(toolbarButton);
      }
    }, 1000);
  }

  private createFloatingChatPanel() {
    this.chatContainer = document.createElement('div');
    this.chatContainer.id = 'univer-chat-panel';
    this.chatContainer.style.cssText = `
      position: fixed; top: 0; right: -350px; width: 350px; height: 100vh;
      background: white; border-left: 1px solid #e0e0e0; box-shadow: -2px 0 10px rgba(0,0,0,0.1);
      z-index: 1000; transition: right 0.3s ease-in-out; overflow: hidden;
      display: flex; flex-direction: column;
    `;

    document.body.appendChild(this.chatContainer);
    this.renderChatContent();
  }

  private toggleChatPanel() {
    if (this.chatContainer) {
      const isVisible = this.chatContainer.style.right === '0px';
      this.chatContainer.style.right = isVisible ? '-350px' : '0px';
    }
  }

  private renderChatContent() {
    if (!this.chatContainer) return;
    render(this.createChatUI(), this.chatContainer);
  }

  private createChatUI() {
    return html`
      <div class="chat-container" style="display: flex; flex-direction: column; height: 100%; background: #fff; overflow: hidden;">
        <div class="chat-header" style="padding: 16px; background: #6c757d; color: white; font-weight: bold; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
          <span>💬 Chat Assistant</span>
          <button @click=${this.closeChatPanel.bind(this)} style="background: none; border: none; color: white; cursor: pointer; font-size: 20px; padding: 4px; border-radius: 4px; transition: background 0.2s;" onmouseover="this.style.background='rgba(255,255,255,0.2)'" onmouseout="this.style.background='none'">×</button>
        </div>

        <div class="chat-messages" style="flex: 1; overflow-y: auto; padding: 16px; background: #f8f9fa;">
          ${this.messages.length === 0 
            ? html`<div style="color: #666; font-style: italic; text-align: center; margin-top: 40px; font-size: 14px;">👋 Welcome to Chat!<br/>Start a conversation...</div>`
            : this.messages.map((msg, index) => html`
                <div class="message ${msg.type}" style="margin: 12px 0; padding: 12px 16px; background: ${msg.type === 'user' ? '#e3f2fd' : '#fff'}; border-radius: 12px; border: 1px solid #e0e0e0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); word-wrap: break-word; animation: slideIn 0.3s ease-out; max-width: 85%; margin-left: ${msg.type === 'user' ? 'auto' : '0'}; margin-right: ${msg.type === 'user' ? '0' : 'auto'};">
                  <div style="font-size: 14px; margin-bottom: 4px; color: #666; font-weight: 500;">${msg.type === 'user' ? '👤 You' : '🤖 AI Assistant'}</div>
                  <div style="font-size: 14px; line-height: 1.4;">${msg.text}</div>
                  <div style="font-size: 11px; color: #999; margin-top: 4px;">${msg.timestamp.toLocaleTimeString('vi-VN')}</div>
                </div>
              `)
          }
          ${this.isTyping ? html`<div class="typing-indicator" style="margin: 12px 0; padding: 12px 16px; background: #fff; border-radius: 12px; border: 1px solid #e0e0e0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); max-width: 85%; font-style: italic; color: #666;">🤖 AI is typing...</div>` : ''}
        </div>
        
        <div class="chat-input-area" style="padding: 16px; border-top: 1px solid #e0e0e0; background: white; display: flex; gap: 12px; align-items: center; box-shadow: 0 -2px 4px rgba(0,0,0,0.1);">
          <input id="chat-input" type="text" placeholder="Type your message..." style="flex: 1; padding: 12px 16px; border: 1px solid #ddd; border-radius: 24px; font-size: 14px; outline: none; transition: border-color 0.2s;" onfocus="this.style.borderColor='#6c757d'" onblur="this.style.borderColor='#ddd'" @keydown=${this.handleKeyDown.bind(this)} />
          <button @click=${this.handleSendMessage.bind(this)} style="padding: 12px 20px; background: #6c757d; color: white; border: none; border-radius: 24px; cursor: pointer; font-size: 14px; font-weight: 500; min-width: 70px; transition: all 0.2s;" onmouseover="this.style.background='#5a6268'; this.style.transform='scale(1.05)'" onmouseout="this.style.background='#6c757d'; this.style.transform='scale(1)'">Send</button>
        </div>
      </div>

      <style>
        @keyframes slideIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
      </style>
    `;
  }

  private closeChatPanel() {
    if (this.chatContainer) {
      this.chatContainer.style.right = '-350px';
    }
  }

  private handleKeyDown(event: KeyboardEvent) {
    if (event.key === 'Enter') {
      event.preventDefault();
      this.handleSendMessage();
    }
  }

  private async handleSendMessage() {
    const input = document.querySelector('#chat-input') as HTMLInputElement;
    
    if (input && input.value.trim()) {
      const userMessage = input.value.trim();
      
      this.addUserMessage(userMessage);
      input.value = "";
      
      this.setTyping(true);
      
      try {
        const response = await this.callMCPService(userMessage);
        this.addBotMessage(response);
      } catch (error) {
        this.addBotMessage("Sorry, I'm experiencing technical issues. Please try again later.");
      } finally {
        this.setTyping(false);
        input.focus();
      }
    }
  }

  private async callMCPService(message: string): Promise<string> {
    if (message.toLowerCase().startsWith('set ai key ')) {
      const aiApiKey = message.substring(11).trim();
      if (aiApiKey) {
        this.aiConfig.aiApiKey = aiApiKey;
        localStorage.setItem('univer_ai_api_key', aiApiKey);
        return "✅ Đã lưu API key! Bạn có thể bắt đầu gõ lệnh.";
      } else {
        return "❌ API key không hợp lệ.";
      }
    }
    if (!this.aiConfig.aiApiKey) {
      return '⚠️ Bạn chưa cấu hình API key. Gõ: set ai key YOUR_KEY_HERE';
    }
    return await this.processWithAI(message);
  }

  private async processWithAI(message: string): Promise<string> {
    try {
      const context = await this.getSpreadsheetContext();
      const aiAnalysis = await this.callAIModel(message, context);
      const results = await this.executeOperations(aiAnalysis.operations);
      
      return `🤖 AI: ${aiAnalysis.explanation}\n\n✅ Results:\n${results.join('\n')}`;
      
    } catch (error) {
      return `❌ Error: ${error instanceof Error ? error.message : 'AI processing failed'}`;
    }
  }

  private async callAIModel(message: string, context: string): Promise<{explanation: string, operations: any[]}> {
    const systemPrompt = `You are an assistant that converts natural language spreadsheet requests into JSON operations for a Univer spreadsheet.
Return ONLY valid JSON (no markdown) matching this schema:
{"explanation":"string","operations":[{"type":"one_of(create_revenue_table,create_employee_table,create_sample_table,sum_column,clear_selection,search_data,set_cell_value)","column?":"A|B|C...","cell?":"e.g. B3","value?":"any","keyword?":"any"}]}

Rules:
- If user wants to insert text into a cell: type=set_cell_value with cell + value
- If user says 'sum column X' -> type=sum_column with column=X (single letter)
- If user asks to clear -> type=clear_selection
- If user wants a table (revenue/employee/sample) -> corresponding create_* table
- If search/find -> type=search_data with keyword
- If ambiguity, default to write text into A1
CurrentContext: ${context}
User: ${message}`;

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.aiConfig.aiApiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-3.5-turbo',
        temperature: 0.1,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: message }
        ]
      })
    });

    if (!response.ok) {
      throw new Error('AI call failed ' + response.status);
    }
    const data = await response.json();
    const raw = data.choices?.[0]?.message?.content || '';
    try {
      // Clean possible code fences if any
      const cleaned = raw.trim().replace(/^```json/,'').replace(/^```/,'').replace(/```$/,'').trim();
      return JSON.parse(cleaned);
    } catch {
      return {
        explanation: `Parsed fallback for: ${message}`,
        operations: [{ type: 'set_cell_value', cell: 'A1', value: message }]
      };
    }
  }

  private async executeOperations(operations: any[]): Promise<string[]> {
    const results: string[] = [];
    
    for (const operation of operations) {
      try {
        const operationType = operation.type || operation.tool;
        
        switch (operationType) {
          case 'create_revenue_table':
            await this.createRevenueTable();
            results.push('✅ Created revenue table');
            break;
            
          case 'create_employee_table':
            await this.createEmployeeTable();
            results.push('✅ Created employee table');
            break;
            
          case 'create_sample_table':
            await this.createSampleTable();
            results.push('✅ Created sample table');
            break;
            
          case 'sum_column':
            const column = operation.column || 'B';
            await this.sumColumn(column);
            results.push(`✅ Summed column ${column}`);
            break;
            
          case 'clear_selection':
            await this.clearSelection();
            results.push('✅ Cleared data');
            break;
            
          case 'search_data':
            const keyword = operation.keyword || operation.arguments?.keyword || 'test';
            const searchResults = await this.searchInSheet(keyword);
            results.push(`🔍 Found ${searchResults} results`);
            break;
            
          case 'set_cell_value':
            const value = operation.value || operation.arguments?.value || 'default';
            const cell = operation.cell || operation.arguments?.cell || 'A1';
            await this.setCellValue(cell, value);
            results.push(`✅ Added data to cell ${cell}`);
            break;
            
          default:
            results.push(`❓ Unknown operation: ${operationType || 'undefined'}`);
        }
      } catch (error) {
        results.push(`❌ Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    }
    
    return results;
  }

  private async createRevenueTable(): Promise<void> {
    const revenueData = [
      ['Month', 'Revenue', 'Profit'],
      ['January', 1000000, 200000],
      ['February', 1200000, 250000],
      ['March', 1100000, 220000],
      ['April', 1300000, 270000],
      ['May', 1150000, 230000],
      ['June', 1400000, 290000]
    ];
    
    await this.setRangeData(revenueData);
  }

  private async createEmployeeTable(): Promise<void> {
    const employeeData = [
      ['Name', 'Position', 'Salary', 'Department'],
      ['John Doe', 'Developer', 15000000, 'IT'],
      ['Jane Smith', 'Designer', 12000000, 'Design'],
      ['Mike Johnson', 'Manager', 20000000, 'Management'],
      ['Sarah Wilson', 'Tester', 10000000, 'QA'],
      ['Tom Brown', 'DevOps', 18000000, 'IT']
    ];
    
    await this.setRangeData(employeeData);
  }

  private async createSampleTable(): Promise<void> {
    const sampleData = [
      ['ID', 'Product', 'Price', 'Quantity'],
      [1, 'Laptop', 15000000, 10],
      [2, 'Mouse', 200000, 50],
      [3, 'Keyboard', 500000, 30],
      [4, 'Monitor', 3000000, 15]
    ];
    
    await this.setRangeData(sampleData);
  }

  private async setRangeData(data: any[][]): Promise<boolean> {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any;
      if (!workbook) return false;

      const worksheet = workbook.getActiveSheet();
      if (!worksheet) return false;

      const valueObject: any = {};
      for (let row = 0; row < data.length; row++) {
        valueObject[row] = {};
        for (let col = 0; col < data[row].length; col++) {
          valueObject[row][col] = { v: data[row][col] };
        }
      }

      const result = await this._commandService.executeCommand(SetRangeValuesCommand.id, {
        unitId: workbook.getUnitId(),
        subUnitId: worksheet.getSheetId(),
        range: {
          startRow: 0,
          endRow: data.length - 1,
          startColumn: 0,
          endColumn: data[0].length - 1,
        },
        value: valueObject,
      });

      return result;
    } catch (error) {
      return false;
    }
  }

  private async sumColumn(column: string): Promise<void> {
    const lastRow = await this.findLastRowInColumn(column);
    if (lastRow > 1) {
      const sumFormula = `=SUM(${column}2:${column}${lastRow})`;
      await this.setCellValue(`${column}${lastRow + 2}`, sumFormula);
    }
  }

  private async findLastRowInColumn(column: string): Promise<number> {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any;
      if (!workbook) return 0;

      const worksheet = workbook.getActiveSheet();
      if (!worksheet) return 0;

      const columnIndex = this.columnToIndex(column);
      let lastRow = 0;

      for (let i = 0; i < 100; i++) {
        const cellData = worksheet.getCell(i, columnIndex);
        if (cellData && cellData.v !== undefined && cellData.v !== null && cellData.v !== '') {
          lastRow = i + 1;
        }
      }

      return lastRow;
    } catch (error) {
      return 0;
    }
  }

  private async clearSelection(): Promise<void> {
    const selection = this.getActiveSelection();
    if (selection) {
      await this.clearRange(selection.startRow, selection.startColumn, selection.endRow, selection.endColumn);
    } else {
      await this.clearRange(0, 0, 9, 9);
    }
  }

  private getActiveSelection(): {startRow: number, startColumn: number, endRow: number, endColumn: number} | null {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any;
      if (!workbook) return null;

      const worksheet = workbook.getActiveSheet();
      if (!worksheet) return null;

      return {
        startRow: 0,
        startColumn: 0,
        endRow: 0,
        endColumn: 0
      };
    } catch (error) {
      return null;
    }
  }

  private async clearRange(startRow: number, startColumn: number, endRow: number, endColumn: number): Promise<boolean> {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any;
      if (!workbook) return false;

      const worksheet = workbook.getActiveSheet();
      if (!worksheet) return false;

      const valueObject: any = {};
      for (let row = startRow; row <= endRow; row++) {
        valueObject[row] = {};
        for (let col = startColumn; col <= endColumn; col++) {
          valueObject[row][col] = { v: '' };
        }
      }

      const result = await this._commandService.executeCommand(SetRangeValuesCommand.id, {
        unitId: workbook.getUnitId(),
        subUnitId: worksheet.getSheetId(),
        range: {
          startRow: startRow,
          endRow: endRow,
          startColumn: startColumn,
          endColumn: endColumn,
        },
        value: valueObject,
      });

      return result;
    } catch (error) {
      return false;
    }
  }

  private async searchInSheet(keyword: string): Promise<number> {
    let count = 0;
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET);
    if (workbook) {
      const worksheet = (workbook as any).getActiveSheet();
      if (worksheet) {
        const range = worksheet.getRange(0, 0, 50, 20);
        const values = range.getValues();
        
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            if (values[i][j] && values[i][j].toString().toLowerCase().includes(keyword.toLowerCase())) {
              count++;
            }
          }
        }
      }
    }
    return count;
  }

  private async getSpreadsheetContext(): Promise<string> {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET);
      if (!workbook) {
        return "No spreadsheet open";
      }

      const worksheet = (workbook as any).getActiveSheet();
      if (!worksheet) {
        return "No active worksheet";
      }

      const context = {
        sheetName: worksheet.getName(),
        totalRows: worksheet.getRowCount(),
        totalCols: worksheet.getColumnCount(),
        hasData: worksheet.getRowCount() > 0
      };

      return JSON.stringify(context);
    } catch (error) {
      return "Unable to get spreadsheet context";
    }
  }

  private async setCellValue(cellAddress: string, value: string): Promise<boolean> {
    try {
      const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any;
      if (!workbook) return false;

      const worksheet = workbook.getActiveSheet();
      if (!worksheet) return false;

      const match = cellAddress.match(/([A-Z]+)(\d+)/);
      if (!match) return false;

      const col = this.columnToIndex(match[1]);
      const row = parseInt(match[2]) - 1;

      const result = await this._commandService.executeCommand(SetRangeValuesCommand.id, {
        unitId: workbook.getUnitId(),
        subUnitId: worksheet.getSheetId(),
        range: {
          startRow: row,
          endRow: row,
          startColumn: col,
          endColumn: col,
        },
        value: {
          [row]: {
            [col]: {
              v: value,
            },
          },
        },
      });

      return result;
    } catch (error) {
      return false;
    }
  }

  private columnToIndex(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
  }

  private addUserMessage(text: string) {
    const message: ChatMessage = {
      text,
      type: 'user',
      timestamp: new Date()
    };
    this.messages.push(message);
    this.renderChatContent();
    this.scrollToBottom();
  }

  private addBotMessage(text: string) {
    const message: ChatMessage = {
      text,
      type: 'bot',
      timestamp: new Date()
    };
    this.messages.push(message);
    this.renderChatContent();
    this.scrollToBottom();
  }

  private setTyping(isTyping: boolean) {
    this.isTyping = isTyping;
    this.renderChatContent();
    if (isTyping) {
      this.scrollToBottom();
    }
  }

  private scrollToBottom() {
    setTimeout(() => {
      const messagesContainer = document.querySelector('.chat-messages');
      if (messagesContainer) {
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
      }
    }, 50);
  }

  public showChat() {
    if (this.chatContainer) {
      this.chatContainer.style.right = '0px';
    }
  }

  public hideChat() {
    if (this.chatContainer) {
      this.chatContainer.style.right = '-350px';
    }
  }

  public toggleChat() {
    if (this.chatContainer) {
      const isVisible = this.chatContainer.style.right === '0px';
      this.chatContainer.style.right = isVisible ? '-350px' : '0px';
    }
  }
}