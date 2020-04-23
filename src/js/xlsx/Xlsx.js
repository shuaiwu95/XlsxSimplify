import * as SheetXlsx from 'xlsx'
import '../../css/table.css'
export default class Xlsx {
  constructor (container = '') {
    this.container = container
    this.sheetLeft_ = 100
    container.classList.add('xlsx-con')
    this.activeIndex = 0
    window.onresize = () => {
      const panels = document.getElementsByClassName('xlsx-tablePanel-con')
      for (let i = 0; i < panels.length; i ++) {
        this.resetTable_(panels[i])
      }
    }
  }

  /**
   * @private
   * @description 获取网络文件流
   * @param {String} url
   * @param {Function} callback
   * @memberof Xlsx
   */
  httpRequest_ (url, callback) {
    if (typeof url === 'string') { // 根据文件路径来展示
      const xhr = new XMLHttpRequest()
      xhr.open('get', url, true)
      xhr.responseType = 'arraybuffer'
      xhr.onload = function() {
        if (xhr.status === 200) {
          const data = new Uint8Array(xhr.response)
          const workbook = SheetXlsx.read(data, {
            'type': 'array',
            'cellStyles': true,
            'sheetStubs': true
          })
          if (callback) {
            callback(workbook)
          }
        } else { // 根据文件流来展示
          const workbook = SheetXlsx.read(url, {'type': 'binary'})
          if (callback) {
            callback(workbook)
          }
        }
      }
      xhr.send()
    } else {
      alert(123)
    }
  }


  /**
   *
   * @description 读取文件
   * @param {string} [url='']
   * @memberof Xlsx
   */
  readFile (url = '') {
    this.httpRequest_(url, this.outputWorkbook_.bind(this))
  }


  /**
   * @private
   * @description 输出文件流
   * @param {Object} workbook
   * @memberof Xlsx
   */
  outputWorkbook_ (workbook) {
    const sheetNames = workbook.SheetNames // 工作表名称集合
    this.container.appendChild(this.createTabMenu_(sheetNames, workbook)) //  创建sheet
  }

  /**
   * @private
   * @description 生成TAB按钮
   * @param {Array} sheetNames
   * @param {Object} workbook
   * @returns
   * @memberof Xlsx
   */
  createTabMenu_ (sheetNames, workbook) {
    // 创建tab
    const tabDiv = document.createElement('div')
    tabDiv.classList.add('xlsx-con-menu')
    const tabDivCon = this.tabDivCon_ = document.createElement('div')
    tabDivCon.classList.add('xlsx-con-menu-con')
    tabDiv.appendChild(tabDivCon)
    sheetNames.forEach((item, index) => {
      const tabItem = document.createElement('span')
      tabItem.classList.add('xlsx-con-menu-item')
      if (index === 0) {
        tabItem.classList.add('xlsx-active')
      }
      tabItem.innerHTML = item
      tabItem.onclick = () => {
        this.clickTab_(index, tabItem)
      }
      tabDivCon.appendChild(tabItem)
      this.createPanel_(item, index, workbook)
    })
    // 创建左右切换
    const clickMenu = document.createElement('div')
    clickMenu.classList.add('xlsx-con-menu-clickMenu-con')
    const zz = this.zz_ = document.createElement('span')
    zz.className = 'iconfont iconzuozuo-'
    // zz.classList.add('xlsx-con-menu-clickMenu-noClick')
    const z = this.z_ = document.createElement('span')
    z.className = 'iconfont iconleft'
    // z.classList.add('xlsx-con-menu-clickMenu-noClick')
    const y = this.y_ = document.createElement('span')
    y.className = 'iconfont iconright'
    const yy = this.yy_ = document.createElement('span')
    yy.className = 'iconfont iconyouyou-'
    clickMenu.appendChild(zz)
    clickMenu.appendChild(z)
    clickMenu.appendChild(y)
    clickMenu.appendChild(yy)
    tabDiv.appendChild(clickMenu)
    this.zz_.onclick = () => {
      this.clickMenu_('zz')
    }
    this.z_.onclick = () => {
      this.clickMenu_('z')
    }
    this.y_.onclick = () => {
      this.clickMenu_('y')
    }
    this.yy_.onclick = () => {
      this.clickMenu_('yy')
    }
    return tabDiv
  }

  /**
   *
   * @description 滑动Tab
   * @param {String} type
   * @memberof Xlsx
   */
  clickMenu_ (type) {
    const containerW = this.container.clientWidth
    const tabContainerW = this.tabDivCon_.clientWidth
    const minLeft = containerW - tabContainerW
    const maxLeft = 100
    switch (type) {
    case 'zz':
      this.sheetLeft_ = 100
      this.tabDivCon_.style.left = `${this.sheetLeft_}px`
      break
    case 'z':
      if (this.sheetLeft_ >= maxLeft) {
        return
      }
      this.sheetLeft_ += 50
      this.tabDivCon_.style.left = `${this.sheetLeft_}px`
      break
    case 'y':
      if (this.sheetLeft_ < minLeft || this.sheetLeft_ === minLeft) {
        return
      }
      this.sheetLeft_ -= 50
      this.tabDivCon_.style.left = `${this.sheetLeft_}px`
      break
    case 'yy':
      if (containerW >= tabContainerW) {
        return
      }
      this.sheetLeft_ = minLeft
      this.tabDivCon_.style.left = `${this.sheetLeft_}px`
      break
    default:
      break
    }
  }

  /**
   * @private
   * @description 为Tab绑定事件
   * @param {Number} index
   * @memberof Xlsx
   */
  clickTab_ (index, tabItem) {
    // alert(index)
    const tabs = document.getElementsByClassName('xlsx-con-menu-item')
    for (let i = 0;i < tabs.length;i ++) {
      tabs[i].classList.remove('xlsx-active')
    }
    tabItem.classList.add('xlsx-active')

    const panels = document.getElementsByClassName('xlsx-tablePanel-con')
    for (let i = 0; i < panels.length; i ++) {
      panels[i].classList.remove('showXlsx-panel')
    }
    panels[index].classList.add('showXlsx-panel')
  }

  /**
   * @private
   * @description 生成Table展示面板
   * @param {String} sheetName
   * @param {Number} index
   * @param {Object} workbook
   * @memberof Xlsx
   */
  createPanel_ (sheetName, index, workbook) {
    const tablePanel = document.createElement('div')
    tablePanel.classList.add('xlsx-tablePanel-con')
    tablePanel.style.display = 'none'
    this.resetTable_(tablePanel)
    if (index === this.activeIndex) {
      tablePanel.classList.add('showXlsx-panel')
    }
    tablePanel.innerHTML = this.getTableHtml_(workbook, sheetName)
    this.container.appendChild(tablePanel)
  }

  /**
   * @private
   * @description 计算表格高度
   * @memberof Xlsx
   */
  resetTable_ (tablePanel) {
    tablePanel.style.height = `${this.container.clientHeight - 30}px`
  }

  /**
   * @private
   * @description 生成表格html
   * @param {Object} workbook
   * @param {String} sheetName
   * @returns
   * @memberof Xlsx
   */
  getTableHtml_ (workbook, sheetName) {
    return SheetXlsx.utils.sheet_to_html(workbook.Sheets[sheetName])
  }
}