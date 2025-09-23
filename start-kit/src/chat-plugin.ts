import { ICommandService, Inject, Injector, IUniverInstanceService, Plugin, UniverInstanceType } from '@univerjs/core'
import { SetRangeValuesCommand } from '@univerjs/sheets'

interface ChatMessage {
  text: string
  type: 'user' | 'bot'
  timestamp: Date
}

interface SimpleAIConfig {
  aiApiKey: string
}

export class ChatPlugin extends Plugin {
  static override pluginName = 'ChatPlugin'

  private messages: ChatMessage[] = []
  private chatContainer: HTMLElement | null = null
  private isTyping = false
  private aiConfig: SimpleAIConfig = { aiApiKey: '' }

  constructor(
    _config: unknown,
    @Inject(Injector) protected readonly _injector: Injector,
    @Inject(ICommandService) private readonly _commandService: ICommandService,
    @Inject(IUniverInstanceService) private readonly _univerInstanceService: IUniverInstanceService,
  ) {
    super()
    this.initializeAIConfig()
  }

  override onReady() {
    this.initializeChatPanel()
  }

  override dispose() {
    if (this.chatContainer)
      this.chatContainer.remove()
  }

  private initializeAIConfig() {
    const savedKey = localStorage.getItem('univer_ai_api_key')
    if (savedKey) {
      this.aiConfig.aiApiKey = savedKey
      setTimeout(() => this.addBotMessage('AI ready. Try: "create revenue table", "sum column B", "put hello in B3"'), 400)
    }
    else {
      setTimeout(() => this.addBotMessage('Enter API key first: type "set ai key YOUR_KEY_HERE"'), 400)
    }
  }

  private initializeChatPanel() {
    this.createFloatingChatPanel()
    this.addToolbarButton()
    ;(window as any).univerChat = {
      show: () => this.showChat(),
      hide: () => this.hideChat(),
      toggle: () => this.toggleChat(),
    }
  }

  private addToolbarButton() {
    const toolbarButton = document.createElement('button')
    toolbarButton.innerHTML = '💬'
    toolbarButton.title = 'Toggle Chat'
    toolbarButton.style.cssText = 'background:none;border:none;cursor:pointer;font-size:18px;padding:8px;margin:0 4px;border-radius:4px'
    toolbarButton.addEventListener('click', () => this.toggleChat())

    setTimeout(() => {
      const toolbar = document.querySelector('.univer-header')
        || document.querySelector('.univer-toolbar')
        || document.querySelector('[data-range-selector="toolbar"]')
        || document.querySelector('.toolbar')
      if (toolbar) {
        toolbar.appendChild(toolbarButton)
      }
      else {
        toolbarButton.style.cssText += 'position:fixed;top:10px;right:10px;z-index:999;background:#f0f0f0;border:1px solid #ccc;border-radius:6px;padding:8px 12px'
        document.body.appendChild(toolbarButton)
      }
    }, 800)
  }

  private createFloatingChatPanel() {
    this.chatContainer = document.createElement('div')
    this.chatContainer.id = 'univer-chat-panel'
    this.chatContainer.style.cssText = 'position:fixed;top:0;right:-350px;width:350px;height:100vh;background:#fff;border-left:1px solid #e0e0e0;box-shadow:-2px 0 10px rgba(0,0,0,.1);z-index:1000;transition:right .3s;overflow:hidden;display:flex;flex-direction:column'
    document.body.appendChild(this.chatContainer)
    this.renderChatContent()
  }

  private renderChatContent() {
    if (!this.chatContainer)
      return
    this.chatContainer.innerHTML = this.createChatHTML()
    const closeBtn = this.chatContainer.querySelector('[data-id="close-btn"]') as HTMLButtonElement | null
    const sendBtn = this.chatContainer.querySelector('[data-id="send-btn"]') as HTMLButtonElement | null
    const input = this.chatContainer.querySelector('#chat-input') as HTMLInputElement | null
    if (closeBtn)
      closeBtn.addEventListener('click', () => this.closeChatPanel())
    if (sendBtn)
      sendBtn.addEventListener('click', () => this.handleSendMessage())
    if (input)
      input.addEventListener('keydown', e => this.handleKeyDown(e as KeyboardEvent))
  }

  private createChatHTML(): string {
    const messagesHTML = this.messages.length === 0
      ? '<div style="color:#666;font-style:italic;text-align:center;margin-top:40px;font-size:14px">Welcome! Start a conversation…</div>'
      : this.messages.map(msg => (
          `<div style="margin:12px 0;padding:12px 16px;background:${msg.type === 'user' ? '#e3f2fd' : '#fff'};border-radius:12px;border:1px solid #e0e0e0;box-shadow:0 1px 3px rgba(0,0,0,.1);max-width:85%;margin-left:${msg.type === 'user' ? 'auto' : '0'};margin-right:${msg.type === 'user' ? '0' : 'auto'}">
           <div style=\"font-size:12px;margin-bottom:4px;color:#666\">${msg.type === 'user' ? 'You' : 'AI'}</div>
           <div style=\"font-size:14px;line-height:1.4\">${this.escapeHtml(msg.text)}</div>
         </div>`
        )).join('')
    const typingHTML = this.isTyping ? '<div style="margin:12px 0;padding:12px 16px;background:#fff;border:1px solid #e0e0e0;border-radius:12px;color:#666;font-style:italic">AI is typing…</div>' : ''
    return `
      <div style="display:flex;flex-direction:column;height:100%;background:#fff;overflow:hidden">
        <div style="padding:16px;background:#6c757d;color:#fff;font-weight:bold;display:flex;justify-content:space-between;align-items:center;box-shadow:0 2px 4px rgba(0,0,0,.1)">
          <span>Chat Assistant</span>
          <button data-id="close-btn" style="background:none;border:none;color:#fff;cursor:pointer;font-size:20px">×</button>
        </div>
        <div class="chat-messages" style="flex:1;overflow-y:auto;padding:16px;background:#f8f9fa">${messagesHTML}${typingHTML}</div>
        <div style="padding:16px;border-top:1px solid #e0e0e0;background:#fff;display:flex;gap:12px;align-items:center">
          <input id="chat-input" type="text" placeholder="Type your message..." style="flex:1;padding:12px 16px;border:1px solid #ddd;border-radius:24px;font-size:14px;outline:none" />
          <button data-id="send-btn" style="padding:12px 20px;background:#6c757d;color:#fff;border:none;border-radius:24px;cursor:pointer;font-size:14px">Send</button>
        </div>
      </div>
    `
  }

  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  }

  private closeChatPanel() {
    if (this.chatContainer)
      this.chatContainer.style.right = '-350px'
  }

  public showChat() {
    if (this.chatContainer)
      this.chatContainer.style.right = '0px'
  }

  public hideChat() {
    if (this.chatContainer)
      this.chatContainer.style.right = '-350px'
  }

  public toggleChat() {
    if (!this.chatContainer)
      return
    const isVisible = this.chatContainer.style.right === '0px'
    this.chatContainer.style.right = isVisible ? '-350px' : '0px'
  }

  private handleKeyDown(event: KeyboardEvent) {
    if (event.key === 'Enter') {
      event.preventDefault()
      this.handleSendMessage()
    }
  }

  private async handleSendMessage() {
    const input = document.querySelector('#chat-input') as HTMLInputElement
    if (input && input.value.trim()) {
      const userMessage = input.value.trim()
      this.addUserMessage(userMessage)
      input.value = ''
      this.setTyping(true)
      try {
        const response = await this.callMCPService(userMessage)
        this.addBotMessage(response)
      }
      catch {
        this.addBotMessage('Sorry, I\'m experiencing technical issues. Please try again later.')
      }
      finally {
        this.setTyping(false)
        input.focus()
      }
    }
  }

  private addUserMessage(text: string) { this.messages.push({ text, type: 'user', timestamp: new Date() }); this.renderChatContent(); this.scrollToBottom() }
  private addBotMessage(text: string) { this.messages.push({ text, type: 'bot', timestamp: new Date() }); this.renderChatContent(); this.scrollToBottom() }
  private setTyping(isTyping: boolean) {
    this.isTyping = isTyping; this.renderChatContent(); if (isTyping)
      this.scrollToBottom()
  }

  private scrollToBottom() {
    setTimeout(() => {
      const el = document.querySelector('.chat-messages') as HTMLElement; if (el)
        el.scrollTop = el.scrollHeight
    }, 50)
  }

  private async callMCPService(message: string): Promise<string> {
    if (message.toLowerCase().startsWith('set ai key ')) {
      const aiApiKey = message.substring(11).trim()
      if (aiApiKey) {
        this.aiConfig.aiApiKey = aiApiKey
        localStorage.setItem('univer_ai_api_key', aiApiKey)
        return 'API key saved. You can start typing commands.'
      }
      return 'Invalid API key.'
    }
    if (!this.aiConfig.aiApiKey)
      return 'No API key configured. Type: set ai key YOUR_KEY_HERE'
    return await this.processWithAI(message)
  }

  private async processWithAI(message: string): Promise<string> {
    try {
      const context = await this.getSpreadsheetContext()
      const aiAnalysis = await this.callAIModel(message, context)
      const results = await this.executeOperations(aiAnalysis.operations)
      return `${aiAnalysis.explanation}\n\n${results.join('\n')}`
    }
    catch (e) {
      return `Error: ${e instanceof Error ? e.message : 'AI processing failed'}`
    }
  }

  private async callAIModel(message: string, context: string): Promise<{ explanation: string, operations: any[] }> {
    const systemPrompt = `You are an assistant that converts natural language spreadsheet requests into JSON operations for a Univer spreadsheet. Return ONLY valid JSON (no markdown) matching this schema:{"explanation":"string","operations":[{"type":"one_of(create_revenue_table,create_employee_table,create_sample_table,sum_column,clear_selection,search_data,set_cell_value)","column?":"A|B|C...","cell?":"e.g. B3","value?":"any","keyword?":"any"}]} Rules: - If user wants to insert text into a cell: type=set_cell_value with cell + value - If user says 'sum column X' -> type=sum_column with column=X (single letter) - If user asks to clear -> type=clear_selection - If user wants a table (revenue/employee/sample) -> corresponding create_* table - If search/find -> type=search_data with keyword - If ambiguity, default to write text into A1 CurrentContext: ${context} User: ${message}`

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.aiConfig.aiApiKey}`,
      },
      body: JSON.stringify({
        model: 'gpt-3.5-turbo',
        temperature: 0.1,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: message },
        ],
      }),
    })
    if (!response.ok)
      throw new Error(`AI call failed ${response.status}`)
    const data = await response.json()
    const raw = data.choices?.[0]?.message?.content || ''
    try {
      const cleaned = raw.trim().replace(/^```json/, '').replace(/^```/, '').replace(/```$/, '').trim()
      return JSON.parse(cleaned)
    }
    catch {
      return { explanation: `Parsed fallback for: ${message}`, operations: [{ type: 'set_cell_value', cell: 'A1', value: message }] }
    }
  }

  private async executeOperations(operations: any[]): Promise<string[]> {
    const results: string[] = []
    for (const op of operations) {
      try {
        const type = op.type || op.tool
        switch (type) {
          case 'create_revenue_table': await this.createRevenueTable(); results.push('Created revenue table'); break
          case 'create_employee_table': await this.createEmployeeTable(); results.push('Created employee table'); break
          case 'create_sample_table': await this.createSampleTable(); results.push('Created sample table'); break
          case 'sum_column': await this.sumColumn(op.column || 'B'); results.push(`Summed column ${op.column || 'B'}`); break
          case 'clear_selection': await this.clearSelection(); results.push('Cleared data'); break
          case 'search_data': results.push(`Found ${await this.searchInSheet(op.keyword || 'test')} results`); break
          case 'set_cell_value': await this.setCellValue(op.cell || 'A1', op.value ?? ''); results.push(`Set ${op.cell || 'A1'}`); break
          default: results.push(`Unknown operation: ${type || 'undefined'}`)
        }
      }
      catch (e) {
        results.push(`Error: ${e instanceof Error ? e.message : 'Unknown error'}`)
      }
    }
    return results
  }

  private async createRevenueTable() {
    await this.setRangeData([
      ['Month', 'Revenue', 'Profit'],
      ['January', 1000000, 200000],
      ['February', 1200000, 250000],
      ['March', 1100000, 220000],
    ])
  }

  private async createEmployeeTable() {
    await this.setRangeData([
      ['Name', 'Position', 'Salary', 'Department'],
      ['John Doe', 'Developer', 15000000, 'IT'],
      ['Jane Smith', 'Designer', 12000000, 'Design'],
    ])
  }

  private async createSampleTable() {
    await this.setRangeData([
      ['ID', 'Product', 'Price', 'Quantity'],
      [1, 'Laptop', 15000000, 10],
      [2, 'Mouse', 200000, 50],
    ])
  }

  private async setRangeData(data: any[][]): Promise<boolean> {
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return false
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return false
    const valueObject: any = {}
    for (let r = 0; r < data.length; r++) { valueObject[r] = {}; for (let c = 0; c < data[r].length; c++) valueObject[r][c] = { v: data[r][c] } }
    return await this._commandService.executeCommand(SetRangeValuesCommand.id, {
      unitId: workbook.getUnitId(),
      subUnitId: worksheet.getSheetId(),
      range: { startRow: 0, endRow: data.length - 1, startColumn: 0, endColumn: data[0].length - 1 },
      value: valueObject,
    })
  }

  private async setCellValue(cellAddress: string, value: string): Promise<boolean> {
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return false
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return false
    const match = cellAddress.match(/([A-Z]+)(\d+)/)
    if (!match)
      return false
    const col = this.columnToIndex(match[1])
    const row = Number.parseInt(match[2]) - 1
    return await this._commandService.executeCommand(SetRangeValuesCommand.id, {
      unitId: workbook.getUnitId(),
      subUnitId: worksheet.getSheetId(),
      range: { startRow: row, endRow: row, startColumn: col, endColumn: col },
      value: { [row]: { [col]: { v: value } } },
    })
  }

  private async sumColumn(column: string): Promise<void> {
    const last = await this.findLastRowInColumn(column)
    if (last > 1)
      await this.setCellValue(`${column}${last + 2}`, `=SUM(${column}2:${column}${last})`)
  }

  private async findLastRowInColumn(column: string): Promise<number> {
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return 0
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return 0
    const colIndex = this.columnToIndex(column)
    let last = 0
    for (let i = 0; i < 100; i++) {
      const cell = worksheet.getCell(i, colIndex); if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '')
        last = i + 1
    }
    return last
  }

  private async clearSelection(): Promise<void> { await this.clearRange(0, 0, 9, 9) }
  private async clearRange(startRow: number, startColumn: number, endRow: number, endColumn: number): Promise<boolean> {
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return false
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return false
    const valueObject: any = {}
    for (let r = startRow; r <= endRow; r++) { valueObject[r] = {}; for (let c = startColumn; c <= endColumn; c++) valueObject[r][c] = { v: '' } }
    return await this._commandService.executeCommand(SetRangeValuesCommand.id, {
      unitId: workbook.getUnitId(),
      subUnitId: worksheet.getSheetId(),
      range: { startRow, endRow, startColumn, endColumn },
      value: valueObject,
    })
  }

  private async searchInSheet(keyword: string): Promise<number> {
    let count = 0
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return 0
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return 0
    const range = worksheet.getRange(0, 0, 50, 20)
    const values = range.getValues()
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (values[i][j] && values[i][j].toString().toLowerCase().includes(keyword.toLowerCase()))
          count++
      }
    }
    return count
  }

  private async getSpreadsheetContext(): Promise<string> {
    const workbook = this._univerInstanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as any
    if (!workbook)
      return 'No spreadsheet open'
    const worksheet = workbook.getActiveSheet()
    if (!worksheet)
      return 'No active worksheet'
    return JSON.stringify({ sheetName: worksheet.getName(), totalRows: worksheet.getRowCount(), totalCols: worksheet.getColumnCount(), hasData: worksheet.getRowCount() > 0 })
  }

  private columnToIndex(column: string): number { let res = 0; for (let i = 0; i < column.length; i++) res = res * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1); return res - 1 }
}
