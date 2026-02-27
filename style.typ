
// 從任意內容節點中遞歸提取純文字。
#let extract-text(content) = {
  if content.has("text") {
    content.text
  } else if content.has("children") {
    content.children.map(extract-text).join("")
  } else if content.has("body") {
    extract-text(content.body)
  } else {
    " "
  }
}

// Function to extract blocks of text (paragraphs, headings)
// 將文檔內容扁平化為可排版的塊列表（標題/段落/文本）。
#let extract-blocks(content) = {
   if content.func() == heading {
      let heading-fields = content.fields()
      let heading-level = heading-fields.at("depth", default: heading-fields.at("level", default: 1))
      ((kind: "heading", level: heading-level, text: extract-text(content)),)
   } else if content.func() == par {
      ((kind: "par", level: 0, text: extract-text(content)),)
  } else if content.has("children") {
    content.children.map(extract-blocks).flatten()
  } else if content.has("text") {
      ((kind: "text", level: 0, text: content.text),)
  } else if content.has("body") {
     extract-blocks(content.body)
  } else {
    ()
  }
}

// 解析文本片段，按【】分離正文與夾註。
#let parse-segments(text) = {
  let segments = ()
  let current-text = ""
  let in-note = false

   for char in text.clusters() {
      if char == "【" {
      if current-text != "" {
        segments.push((type: if in-note { "note" } else { "text" }, content: current-text))
        current-text = ""
      }
      in-note = true
   } else if char == "】" {
      if current-text != "" {
        segments.push((type: if in-note { "note" } else { "text" }, content: current-text))
        current-text = ""
      }
      in-note = false
    } else {
      current-text += char
    }
  }
  
  if current-text != "" {
    segments.push((type: if in-note { "note" } else { "text" }, content: current-text))
  }
  segments
}

// 內部占位符：佔一字格但不渲染字形（U+2060 Word Joiner，僅供模板內部使用）。
#let grid-space-marker = "\u{2060}"

// 判斷字符是否為不佔格標點。
#let is-punct(char) = {
   (
      "，", "。", "、", "；", "：", "？", "！", "「", "」", "『", "』", "【", "】",
      "（", "）", "《", "》", "〈", "〉", "—", "…", ",", ".", ";", ":", "?", "!"
   ).contains(char)
}

// 將 0-9 轉為中文數字。
#let chinese-digit(n) = {
   if n == 0 { "零" }
   else if n == 1 { "一" }
   else if n == 2 { "二" }
   else if n == 3 { "三" }
   else if n == 4 { "四" }
   else if n == 5 { "五" }
   else if n == 6 { "六" }
   else if n == 7 { "七" }
   else if n == 8 { "八" }
   else { "九" }
}

// 將整數轉為中文數字（目前覆蓋到千位）。
#let to-chinese-number(n) = {
   if n < 10 {
      chinese-digit(n)
   } else if n < 20 {
      "十" + if n == 10 { "" } else { chinese-digit(n - 10) }
   } else if n < 100 {
      let tens = 0
      let rem = n
      while rem >= 10 {
         rem -= 10
         tens += 1
      }
      chinese-digit(tens) + "十" + if rem == 0 { "" } else { chinese-digit(rem) }
   } else if n < 1000 {
      let hundreds = 0
      let rem = n
      while rem >= 100 {
         rem -= 100
         hundreds += 1
      }
      chinese-digit(hundreds) + "百" + if rem == 0 {
         ""
      } else if rem < 10 {
         "零" + chinese-digit(rem)
      } else {
         to-chinese-number(rem)
      }
   } else {
      str(n)
   }
}

// 渲染版心（頁碼、題名、魚尾與中縫）。
#let render-banxin(
   page-idx,
   banxin-title-map,
   page-width,
   middle-col,
   cell-width,
   cell-height,
   char-per-col,
   font-size,
   note-font-size,
   v-shift: 2pt,
   toc-pages: 0,
) = {
   let marks = ()

   // 頁碼：中間空列、倒數第4格（夾註樣式）
   let logical-page = page-idx - toc-pages + 1
   let show-page-number = logical-page > 0
   let page-number = if show-page-number { to-chinese-number(logical-page) } else { "" }
   let mark-x = page-width - (middle-col + 1) * cell-width
   let mark-row = char-per-col - 5
   let mark-y = mark-row * cell-height
   let mark-sub-w = cell-width / 2
   let mark-sub-h = cell-height / 2
   let mark-chars = page-number.clusters()

   let mark-right-count = calc.min(mark-chars.len(), 2 * calc.ceil(mark-chars.len() / 4))

   let mark-idx = 0
   while show-page-number and mark-idx < mark-chars.len() and mark-idx < 4 {
      let in-right = mark-idx < mark-right-count
      let local = if in-right { mark-idx } else { mark-idx - mark-right-count }
      let dy = if calc.rem(local, 2) == 1 { mark-sub-h } else { 0pt }
      let dx = if in-right { mark-sub-w } else { 0pt }

      marks.push(place(
         top + left,
         dx: mark-x + dx,
         dy: mark-y + dy,
         block(
            width: mark-sub-w,
            height: mark-sub-h,
            align(center + horizon, text(size: note-font-size, mark-chars.at(mark-idx)))
         )
      ))
      mark-idx += 1
   }

   // 版心標題：優先二級（去夾註），缺省回退一級（去夾註）
   let page-level-title = ""
   for item in banxin-title-map {
      if page-idx >= item.start {
         page-level-title = item.title
      }
   }

   let title-start-row = 4
   let title-end-row = mark-row
   let title-chars = page-level-title.clusters()
   let title-idx = 0
   for title-row in range(title-start-row, title-end-row) {
      if title-idx >= title-chars.len() {
         break
      }
      let tchar = title-chars.at(title-idx)
      marks.push(place(
         top + left,
         dx: mark-x,
         dy: title-row * cell-height,
         block(
            width: cell-width,
            height: cell-height,
            align(center + horizon, text(size: font-size, tchar))
         )
      ))
      title-idx += 1
   }

   // 版心美化：魚尾與中縫
   let seam-stroke = 2pt
   let fish-stroke = 1.0pt
   let center-x = mark-x + cell-width / 2
   let fish-left = mark-x + cell-width * 0.2
   let fish-right = mark-x + cell-width * 0.8

   // 上魚尾（開口向下）
   let upper-top-y = title-start-row * cell-height - cell-height * 0.3
   let upper-mid-y = title-start-row * cell-height - cell-height * 0.1
   let upper-bottom-y = title-start-row * cell-height + cell-height * 0.1
   marks.push(place(
    top+left,
    dx: 0pt,
    dy: 0pt,
   polygon(
    fill: black,
    (mark-x,upper-top-y),
    (mark-x + cell-width,upper-top-y),
    (mark-x + cell-width,upper-bottom-y),
    (center-x,upper-mid-y),
    (mark-x,upper-bottom-y),
    (mark-x,upper-top-y),
   )))
   

   // 中縫：最上橫線連到上魚尾
   marks.push(place(
      top + left,
      dx: center-x,
      dy: 0pt,
      line(start: (0pt, -v-shift), end: (0pt, upper-top-y), stroke: seam-stroke)
   ))

   // 下魚尾（開口向上）
   let upper-top-y = (mark-row + 1) * cell-height + cell-height * 0.3
   let upper-mid-y = (mark-row + 1) * cell-height + cell-height * 0.1
   let upper-bottom-y = (mark-row + 1) * cell-height - cell-height * 0.1
   marks.push(place(
    top+left,
    dx: 0pt,
    dy: 0pt,
    polygon(
      fill: black,
      (mark-x,upper-top-y),
      (mark-x + cell-width,upper-top-y),
      (mark-x + cell-width,upper-bottom-y),
      (center-x,upper-mid-y),
      (mark-x,upper-bottom-y),
      (mark-x,upper-top-y),
    )
   ))

   marks.push(place(
      top + left,
      dx: center-x,
      dy: 0pt,
      line(start: (0pt, cell-height*char-per-col + v-shift), end: (0pt, upper-top-y), stroke: seam-stroke)
   ))

   marks.join()
}

// 去除標題中的夾註與空格佔位符，只保留正文字符。
#let strip-heading-notes(text) = {
   parse-segments(text)
      .filter(seg => seg.type == "text")
      .map(seg => seg.content.replace(grid-space-marker, ""))
      .join("")
}

// 計算夾註首小列（先填列）字數。
#let note-first-column-count(chars-count, grids-used) = {
   calc.min(chars-count, 2 * grids-used)
}

// 在同頁內前進一列，必要時換頁到首列。
#let next-col(page, col, cols-per-page) = {
   let npage = page
   let ncol = col
   if ncol + 1 < cols-per-page {
      ncol += 1
   } else {
      ncol = 0
      npage += 1
   }
   (page: npage, col: ncol)
}

// 前進到下一列並將行指標重置為 0。
#let move-to-next-col(page, col, cols-per-page) = {
   let stepped = next-col(page, col, cols-per-page)
   (page: stepped.page, col: stepped.col, row: 0)
}

// 行越界時自動換列（並可能換頁）。
#let normalize-row(page, col, row, char-per-col, cols-per-page) = {
   if row >= char-per-col {
      move-to-next-col(page, col, cols-per-page)
   } else {
      (page: page, col: col, row: row)
   }
}

// 行指標前進指定步數，並做越界歸一化。
#let advance-row(page, col, row, char-per-col, cols-per-page, steps: 1) = {
   let state = normalize-row(page, col, row, char-per-col, cols-per-page)
   state = (page: state.page, col: state.col, row: state.row + steps)
   normalize-row(state.page, state.col, state.row, char-per-col, cols-per-page)
}

// 為二/三級標題預留「空一列」規則所需位置。
#let reserve-heading-column(page, col, row, cols-per-page) = {
   let state = (page: page, col: col, row: row)
   if state.row > 0 {
      state = move-to-next-col(state.page, state.col, cols-per-page)
   }
   move-to-next-col(state.page, state.col, cols-per-page)
}

// 區塊結束時收束到下一列起始行。
#let flush-block-end(page, col, row, cols-per-page) = {
   if row > 0 {
      move-to-next-col(page, col, cols-per-page)
   } else {
      (page: page, col: col, row: row)
   }
}

// 確保頁面內容容器至少包含指定頁索引。
#let ensure-page-content(pages-content, page-idx) = {
   let out = pages-content
   while out.len() <= page-idx {
      out.push(())
   }
   out
}

// 根據排版參數計算頁面幾何與字號配置。
#let build-page-layout(char-per-col, cols-per-half-page, page-width: 29.7cm, x-margin: 3cm) = {
   let cols-per-page = 2 * cols-per-half-page
   let total-cols = cols-per-page + 1
   let middle-col = calc.floor(total-cols / 2)
   let cell-width = (page-width - 2 * x-margin) / total-cols
   let cell-height = cell-width / 1.5
   let page-height = cell-height * char-per-col
   let font-size = cell-height
   let note-font-size = font-size * 0.5
   let punct-font-size = font-size * 0.45

   (
      page-width: page-width,
      x-margin: x-margin,
      cols-per-page: cols-per-page,
      total-cols: total-cols,
      middle-col: middle-col,
      cell-width: cell-width,
      cell-height: cell-height,
      page-height: page-height,
      font-size: font-size,
      note-font-size: note-font-size,
      punct-font-size: punct-font-size,
   )
}

// 由標題列表構造目錄塊序列（縮進以空格佔位符嵌入文本）。
#let build-toc-blocks(toc-items, page-offset: 0) = {
   let toc = ((kind: "heading", level: 1, text: "目錄"),)
   for item in toc-items {
      let indent = if item.level == 1 { 0 } else if item.level == 2 { 1 } else { 2 }
      let page-text = to-chinese-number(item.page + 1)
      let toc-title = strip-heading-notes(item.title)
      let prefix = range(indent).fold("", (s, _) => s + grid-space-marker)
      toc.push((
         kind: "toc",
         level: 0,
         text: prefix + toc-title + "【" + page-text + "】",
      ))
   }
   toc
}

// 在指定頁面格位放置一個正文字符。
#let place-cell-char(pages-content, page-idx, x, y, cell-width, cell-height, font-size, char) = {
   let out = ensure-page-content(pages-content, page-idx)
   out.at(page-idx).push(place(
      top + left,
      dx: x,
      dy: y,
      block(width: cell-width, height: cell-height, align(center + horizon, text(size: font-size, char)))
   ))
   out
}

// 將標點壓置到錨點字符右下角。
#let place-anchor-punct(pages-content, anchor, punct-font-size, char) = {
   let out = ensure-page-content(pages-content, anchor.page)
   out.at(anchor.page).push(place(
      top + left,
      dx: anchor.x + anchor.w * 0.62,
      dy: anchor.y + anchor.h * 0.62,
      block(
         width: anchor.w * 0.38,
         height: anchor.h * 0.38,
         align(center + horizon, text(size: punct-font-size, fill: red, char))
      )
   ))
   out
}

// 僅模擬走位，用於估算頁數與標題頁映射。
#let simulate-layout(blocks, char-per-col, cols-per-page) = {
   let current-page = 0
   let current-col = 0
   let current-row = 0
   let headings = ()
   let seen-level1-heading = false

   for blk in blocks {
      if blk.kind == "heading" and blk.level == 1 {
         if seen-level1-heading {
            current-row = 0
            current-col = 0
            current-page += 1
         }
         headings.push((level: 1, title: blk.text, page: current-page))
         seen-level1-heading = true
      }

      if blk.kind == "heading" and blk.level == 2 {
         let state = reserve-heading-column(current-page, current-col, current-row, cols-per-page)
         current-page = state.page
         current-col = state.col
         current-row = state.row
         headings.push((level: 2, title: blk.text, page: current-page))
      }

      if blk.kind == "heading" and blk.level == 3 {
         let state = reserve-heading-column(current-page, current-col, current-row, cols-per-page)
         current-page = state.page
         current-col = state.col
         current-row = state.row
         headings.push((level: 3, title: blk.text, page: current-page))
      }

      let segments = parse-segments(blk.text)
      for segment in segments {
         if segment.type == "text" {
            for char in segment.content.clusters() {
               if not is-punct(char) {
                  let state = advance-row(current-page, current-col, current-row, char-per-col, cols-per-page)
                  current-page = state.page
                  current-col = state.col
                  current-row = state.row
               }
            }
         } else {
            let total-chars = 0
            for c in segment.content.clusters() {
               if not is-punct(c) {
                  total-chars += 1
               }
            }

            let grids-total = calc.ceil(total-chars / 4)
            let grids-processed = 0

            while grids-processed < grids-total {
               let state = normalize-row(current-page, current-col, current-row, char-per-col, cols-per-page)
               current-page = state.page
               current-col = state.col
               current-row = state.row

               let rows-left = char-per-col - current-row
               let grids-in-segment = calc.min(rows-left, grids-total - grids-processed)

               grids-processed += grids-in-segment
               state = advance-row(current-page, current-col, current-row, char-per-col, cols-per-page, steps: grids-in-segment)
               current-page = state.page
               current-col = state.col
               current-row = state.row
            }
         }
      }

      let state = flush-block-end(current-page, current-col, current-row, cols-per-page)
      current-page = state.page
      current-col = state.col
      current-row = state.row
   }

   (pages: current-page + 1, headings: headings)
}

// 迭代收斂目錄頁數與目錄塊內容。
#let resolve-toc-layout(toc-items, char-per-col, cols-per-page, max-iter: 3) = {
   let toc-pages = 0
   let toc-blocks = build-toc-blocks(toc-items, page-offset: 0)
   let iter = 0

   while iter < max-iter {
      toc-pages = simulate-layout(toc-blocks, char-per-col, cols-per-page).pages
      let next-toc = build-toc-blocks(toc-items, page-offset: toc-pages)
      if next-toc == toc-blocks {
         break
      }
      toc-blocks = next-toc
      iter += 1
   }

   (toc-pages: toc-pages, toc-blocks: toc-blocks)
}

// 對單個塊應用標題前置規則與目錄縮進規則。
#let apply-block-prefix(
   blk,
   current-page,
   current-col,
   current-row,
   cols-per-page,
   cols-per-half-page,
   char-per-col,
   seen-level1-heading,
   banxin-title-map,
) = {
   let page = current-page
   let col = current-col
   let row = current-row
   let seen-level1 = seen-level1-heading
   let banxin-map = banxin-title-map

   if blk.kind == "heading" and blk.level == 1 {
      if seen-level1 {
         if row != 0 or col != 0 {
            row = 0
            col = 0
            page += 1
         }
      }
      let start-page = if col < cols-per-half-page { page } else { page + 1 }
      banxin-map.push((start: start-page, title: strip-heading-notes(blk.text), level: 1))
      seen-level1 = true
   }

   if blk.kind == "heading" and blk.level == 2 {
      let state = reserve-heading-column(page, col, row, cols-per-page)
      page = state.page
      col = state.col
      row = state.row
      let start-page = if col < cols-per-half-page { page } else { page + 1 }
      banxin-map.push((start: start-page, title: strip-heading-notes(blk.text), level: 2))
   }

   if blk.kind == "heading" and blk.level == 3 {
      let state = reserve-heading-column(page, col, row, cols-per-page)
      page = state.page
      col = state.col
      row = state.row
   }

   (
      current-page: page,
      current-col: col,
      current-row: row,
      seen-level1-heading: seen-level1,
      banxin-title-map: banxin-map,
   )
}

// 佈局單個塊的正文/夾註片段並更新錨點狀態。
#let layout-block-segments(
   segments,
   col-x,
   cell-width,
   cell-height,
   font-size,
   note-font-size,
   punct-font-size,
   char-per-col,
   cols-per-page,
   current-page,
   current-col,
   current-row,
   pages-content,
   has-last-anchor,
   last-anchor-page,
   last-anchor-x,
   last-anchor-y,
   last-anchor-w,
   last-anchor-h,
   show-punct: true,
) = {
   let page = current-page
   let col = current-col
   let row = current-row
   let content = pages-content
   let has-anchor = has-last-anchor
   let anchor-page = last-anchor-page
   let anchor-x = last-anchor-x
   let anchor-y = last-anchor-y
   let anchor-w = last-anchor-w
   let anchor-h = last-anchor-h

   for segment in segments {
      if segment.type == "text" {
         for char in segment.content.clusters() {
            if is-punct(char) {
               if show-punct and has-anchor {
                  content = place-anchor-punct(
                     content,
                     (page: anchor-page, x: anchor-x, y: anchor-y, w: anchor-w, h: anchor-h),
                     punct-font-size,
                     char,
                  )
               }
            } else if char == grid-space-marker {
               let state = advance-row(page, col, row, char-per-col, cols-per-page)
               page = state.page
               col = state.col
               row = state.row
            } else {
               let x-pos = col-x(col)
               let y-pos = row * cell-height

               content = place-cell-char(
                  content,
                  page,
                  x-pos,
                  y-pos,
                  cell-width,
                  cell-height,
                  font-size,
                  char,
               )

               has-anchor = true
               anchor-page = page
               anchor-x = x-pos
               anchor-y = y-pos
               anchor-w = cell-width
               anchor-h = cell-height

               let state = advance-row(page, col, row, char-per-col, cols-per-page)
               page = state.page
               col = state.col
               row = state.row
            }
         }

      } else {
         let pre-note-has-anchor = has-anchor
         let pre-note-anchor-page = anchor-page
         let pre-note-anchor-x = anchor-x
         let pre-note-anchor-y = anchor-y
         let pre-note-anchor-w = anchor-w
         let pre-note-anchor-h = anchor-h

         let note-chars = ()
         let note-punct = ()
         let logical-char-count = 0
         for c in segment.content.clusters() {
            if is-punct(c) {
               note-punct.push((char: c, target: logical-char-count - 1))
            } else {
               note-chars.push(c)
               logical-char-count += 1
            }
         }

         let total-chars = note-chars.len()
         let grids-total = calc.ceil(total-chars / 4)

         let note-char-anchors = ()
         let chars-processed = 0
         let grids-processed = 0

         while grids-processed < grids-total {
            let state = normalize-row(page, col, row, char-per-col, cols-per-page)
            page = state.page
            col = state.col
            row = state.row

            let rows-left = char-per-col - row
            let grids-remaining = grids-total - grids-processed
            let grids-in-segment = calc.min(rows-left, grids-remaining)

            let capacity = grids-in-segment * 4
            let chars-count = calc.min(total-chars - chars-processed, capacity)
            let segment-chars = note-chars.slice(chars-processed, chars-processed + chars-count)

               let right-count = note-first-column-count(chars-count, grids-in-segment)

            content = ensure-page-content(content, page)

            let base-x = col-x(col)
            let base-y = row * cell-height
            let sub-w = cell-width / 2
            let sub-h = cell-height / 2

            let idx = 0
            while idx < chars-count {
               let in-right = idx < right-count
               let local = if in-right { idx } else { idx - right-count }

               let grid-offset = 0
               let rem2 = local
               while rem2 >= 2 {
                  rem2 -= 2
                  grid-offset += 1
               }

               let char-x = base-x + if in-right { sub-w } else { 0pt }
               let char-y = base-y + grid-offset * cell-height + if rem2 == 1 { sub-h } else { 0pt }

               content.at(page).push(place(
                  top + left,
                  dx: char-x,
                  dy: char-y,
                  block(width: sub-w, height: sub-h, align(center + horizon, text(size: note-font-size, segment-chars.at(idx))))
               ))

               note-char-anchors.push((page: page, x: char-x, y: char-y, w: sub-w, h: sub-h))
               has-anchor = true
               anchor-page = page
               anchor-x = char-x
               anchor-y = char-y
               anchor-w = sub-w
               anchor-h = sub-h

               idx += 1
            }

            chars-processed += chars-count
            grids-processed += grids-in-segment
            state = advance-row(page, col, row, char-per-col, cols-per-page, steps: grids-in-segment)
            page = state.page
            col = state.col
            row = state.row
         }

         for p in note-punct {
            let has-p-anchor = false
            let p-anchor-page = 0
            let p-anchor-x = 0pt
            let p-anchor-y = 0pt
            let p-anchor-w = cell-width
            let p-anchor-h = cell-height

            if p.target >= 0 and p.target < note-char-anchors.len() {
               let a = note-char-anchors.at(p.target)
               has-p-anchor = true
               p-anchor-page = a.page
               p-anchor-x = a.x
               p-anchor-y = a.y
               p-anchor-w = a.w
               p-anchor-h = a.h
            } else if pre-note-has-anchor {
               has-p-anchor = true
               p-anchor-page = pre-note-anchor-page
               p-anchor-x = pre-note-anchor-x
               p-anchor-y = pre-note-anchor-y
               p-anchor-w = pre-note-anchor-w
               p-anchor-h = pre-note-anchor-h
            }

            if show-punct and has-p-anchor {
               content = place-anchor-punct(
                  content,
                  (page: p-anchor-page, x: p-anchor-x, y: p-anchor-y, w: p-anchor-w, h: p-anchor-h),
                  punct-font-size,
                  p.char,
               )
            }
         }
      }
   }

   (
      current-page: page,
      current-col: col,
      current-row: row,
      pages-content: content,
      has-last-anchor: has-anchor,
      last-anchor-page: anchor-page,
      last-anchor-x: anchor-x,
      last-anchor-y: anchor-y,
      last-anchor-w: anchor-w,
      last-anchor-h: anchor-h,
   )
}

// 生成單頁的格線與外框裝飾。
#let build-page-grid-and-frame(page-width, total-cols, cell-width, page-height, v-shift: 2pt, frame-offset: 5pt) = {
   let grid-lines = ()
   let frame-lines = ()

   let grid-left = page-width - total-cols * cell-width
   let grid-right = page-width
   let frame-left = grid-left - frame-offset
   let frame-right = grid-right + frame-offset

   for col-idx in range(total-cols + 1) {
      let x-pos = page-width - (col-idx) * cell-width
      grid-lines.push(place(
         top + left,
         dx: x-pos,
         dy: 0pt,
         line(start: (0pt, -v-shift), end: (0pt, page-height + v-shift), stroke: 0.5pt )
      ))
   }

   frame-lines.push(place(
      top + left,
      dx: frame-left,
      dy: 0pt,
      polygon((0pt, -v-shift), (frame-right - frame-left, -v-shift), (frame-right - frame-left, page-height + v-shift), (0pt, page-height + v-shift), stroke: 2pt)
   ))

   (grid-lines: grid-lines, frame-lines: frame-lines)
}

// 將所有內容塊排版成頁面內容，並產出版心標題映射。
#let layout-all-blocks(
   blocks,
   col-x,
   cols-per-page,
   cols-per-half-page,
   char-per-col,
   cell-width,
   cell-height,
   font-size,
   note-font-size,
   punct-font-size,
   show-punct: true,
) = {
   let current-page = 0
   let current-col = 0
   let current-row = 0

   let has-last-anchor = false
   let last-anchor-page = 0
   let last-anchor-x = 0pt
   let last-anchor-y = 0pt
   let last-anchor-w = cell-width
   let last-anchor-h = cell-height

   let pages-content = ((),)
   let banxin-title-map = ()
   let seen-level1-heading = false

   for blk in blocks {
      let prefix-state = apply-block-prefix(
         blk,
         current-page,
         current-col,
         current-row,
         cols-per-page,
         cols-per-half-page,
         char-per-col,
         seen-level1-heading,
         banxin-title-map,
      )
      current-page = prefix-state.current-page
      current-col = prefix-state.current-col
      current-row = prefix-state.current-row
      seen-level1-heading = prefix-state.seen-level1-heading
      banxin-title-map = prefix-state.banxin-title-map

      let segments = parse-segments(blk.text)
      let segment-state = layout-block-segments(
         segments,
         col-x,
         cell-width,
         cell-height,
         font-size,
         note-font-size,
         punct-font-size,
         char-per-col,
         cols-per-page,
         current-page,
         current-col,
         current-row,
         pages-content,
         has-last-anchor,
         last-anchor-page,
         last-anchor-x,
         last-anchor-y,
         last-anchor-w,
         last-anchor-h,
         show-punct: show-punct,
      )
      current-page = segment-state.current-page
      current-col = segment-state.current-col
      current-row = segment-state.current-row
      pages-content = segment-state.pages-content
      has-last-anchor = segment-state.has-last-anchor
      last-anchor-page = segment-state.last-anchor-page
      last-anchor-x = segment-state.last-anchor-x
      last-anchor-y = segment-state.last-anchor-y
      last-anchor-w = segment-state.last-anchor-w
      last-anchor-h = segment-state.last-anchor-h

      let state = flush-block-end(current-page, current-col, current-row, cols-per-page)
      current-page = state.page
      current-col = state.col
      current-row = state.row
   }

   (pages-content: pages-content, banxin-title-map: banxin-title-map)
}

// 將已排好的頁面內容與裝飾、版心合成最終頁面輸出。
#let render-all-pages(
   pages-content,
   banxin-title-map,
   page-width,
   total-cols,
   middle-col,
   cell-width,
   cell-height,
   page-height,
   char-per-col,
   font-size,
   note-font-size,
   toc-pages,
) = {
   for (page-idx, content) in pages-content.enumerate() {
      let page-deco = build-page-grid-and-frame(page-width, total-cols, cell-width, page-height)
      let grid-lines = page-deco.grid-lines
      let frame-lines = page-deco.frame-lines

      let banxin = render-banxin(
         page-idx,
         banxin-title-map,
         page-width,
         middle-col,
         cell-width,
         cell-height,
         char-per-col,
         font-size,
         note-font-size,
         toc-pages: toc-pages,
      )

      box(width: 100%, height: 100%, { grid-lines.join(); frame-lines.join(); content.join(); banxin })
      if page-idx < pages-content.len() - 1 { pagebreak(weak: true) }
   }
}

// 文檔模板入口：完成抽取、模擬、排版與輸出。
#let template(doc, char-per-col: 20, cols-per-half-page: 10, show-punct: true) = {
   let layout = build-page-layout(char-per-col, cols-per-half-page)
   let page-width = layout.page-width
   let cols-per-page = layout.cols-per-page
   let total-cols = layout.total-cols
   let middle-col = layout.middle-col
   let x-margin = layout.x-margin
   let cell-width = layout.cell-width
   let cell-height = layout.cell-height
   let page-height = layout.page-height
   let font-size = layout.font-size
   let note-font-size = layout.note-font-size
   let punct-font-size = layout.punct-font-size

  
  set page(paper: "a4", flipped: true, margin: (x: -x-margin, top: 4cm), fill: rgb(236,231,184,40%), header-ascent: 0%, footer-descent: 0pt)
  set text(font: "KingHwa_OldSong", size: 16pt)
  // set text(font: "STFangsong", size: 16pt)
  
   let content-blocks = extract-blocks(doc)
      .map(b => (..b, text: b.text.replace(regex("[ \n\t　]"), "")))
      .filter(b => b.text.len() > 0)
      .map(b => if b.kind == "heading" and b.level == 3 {
         // 三級標題前空兩格：以空格佔位符嵌入文本，由排版循環統一處理
         (..b, text: grid-space-marker + grid-space-marker + b.text)
      } else { b })


   let content-sim = simulate-layout(content-blocks, char-per-col, cols-per-page)
   let toc-items = content-sim.headings.filter(h => h.level <= 3)

   let toc-layout = resolve-toc-layout(toc-items, char-per-col, cols-per-page)
   let toc-pages = toc-layout.toc-pages
   let toc-blocks = toc-layout.toc-blocks

   let blocks = toc-blocks + content-blocks

   let logical-col-to-physical = logical-col => if logical-col < middle-col { logical-col } else { logical-col + 1 }
   let col-x = logical-col => page-width - (logical-col-to-physical(logical-col) + 1) * cell-width
    
   let layout-result = layout-all-blocks(
      blocks,
      col-x,
      cols-per-page,
      cols-per-half-page,
      char-per-col,
      cell-width,
      cell-height,
      font-size,
      note-font-size,
      punct-font-size,
      show-punct: show-punct,
   )

   let pages-content = layout-result.pages-content
   let banxin-title-map = layout-result.banxin-title-map

   render-all-pages(
      pages-content,
      banxin-title-map,
      page-width,
      total-cols,
      middle-col,
      cell-width,
      cell-height,
      page-height,
      char-per-col,
      font-size,
      note-font-size,
      toc-pages,
   )
}
