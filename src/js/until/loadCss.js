const scriptPath = 'Xlsx.js'
export default () => {
  document.write(`<link rel="stylesheet" href="${getScriptLocation()}css/main-css.css">`)
}
export const getScriptLocation = (() => {
  let r = new RegExp('(^|(.*?\\/))(' + scriptPath + ')(\\?|$)'),
    s = document.getElementsByTagName('script'),
    src, l = ''
  for (let i=0, len=s.length; i<len; i++) {
    src = s[i].getAttribute('src')
    if (src) {
      let m = src.match(r)
      if (m) {
        l = m[1]
        break
      }
    }
  }
  return function() {
    return l
  }
})()