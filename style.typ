

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

#let is-punct(char) = {
   (
      "，", "。", "、", "；", "：", "？", "！", "「", "」", "『", "』", "【", "】",
      "（", "）", "《", "》", "〈", "〉", "—", "…", ",", ".", ";", ":", "?", "!"
   ).contains(char)
}

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
) = {
   let marks = ()

   // 頁碼：中間空列、倒數第4格（夾註樣式）
   let page-number = to-chinese-number(page-idx + 1)
   let mark-x = page-width - (middle-col + 1) * cell-width
   let mark-row = char-per-col - 5
   let mark-y = mark-row * cell-height
   let mark-sub-w = cell-width / 2
   let mark-sub-h = cell-height / 2
   let mark-chars = page-number.clusters()

   let mark-right-count = if calc.rem(mark-chars.len(), 4) == 1 {
      let k = 0
      let rem = mark-chars.len()
      while rem >= 4 {
         rem -= 4
         k += 1
      }
      2 * k + 1
   } else if calc.rem(mark-chars.len(), 4) == 2 {
      calc.floor(mark-chars.len() / 2)
   } else {
      calc.min(mark-chars.len(), 2)
   }

   let mark-idx = 0
   while mark-idx < mark-chars.len() and mark-idx < 4 {
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

#let strip-heading-notes(text) = {
   parse-segments(text)
      .filter(seg => seg.type == "text")
      .map(seg => seg.content)
      .join("")
}

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
         if current-row > 0 {
            current-row = 0
            if current-col + 1 < cols-per-page {
               current-col += 1
            } else {
               current-col = 0
               current-page += 1
            }
         }
         if current-col + 1 < cols-per-page {
            current-col += 1
         } else {
            current-col = 0
            current-page += 1
         }
         headings.push((level: 2, title: blk.text, page: current-page))
      }

      if blk.kind == "heading" and blk.level == 3 {
         if current-row > 0 {
            current-row = 0
            if current-col + 1 < cols-per-page {
               current-col += 1
            } else {
               current-col = 0
               current-page += 1
            }
         }
         if current-col + 1 < cols-per-page {
            current-col += 1
         } else {
            current-col = 0
            current-page += 1
         }
         if char-per-col > 2 {
            current-row = 2
         } else if char-per-col > 1 {
            current-row = 1
         }
         headings.push((level: 3, title: blk.text, page: current-page))
      }

      if blk.at("toc-indent", default: 0) > 0 {
         let toc-indent = blk.at("toc-indent", default: 0)
         if toc-indent < char-per-col {
            current-row = toc-indent
         } else {
            current-row = char-per-col - 1
         }
      }

      let segments = parse-segments(blk.text)
      for segment in segments {
         if segment.type == "text" {
            for char in segment.content.clusters() {
               if not is-punct(char) {
                  current-row += 1
                  if current-row >= char-per-col {
                     current-row = 0
                     current-col += 1
                     if current-col >= cols-per-page {
                        current-col = 0
                        current-page += 1
                     }
                  }
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
               if current-row >= char-per-col {
                  current-row = 0
                  current-col += 1
                  if current-col >= cols-per-page {
                     current-col = 0
                     current-page += 1
                  }
               }

               let rows-left = char-per-col - current-row
               let grids-in-segment = calc.min(rows-left, grids-total - grids-processed)

               grids-processed += grids-in-segment
               current-row += grids-in-segment

               if current-row >= char-per-col {
                  current-row = 0
                  current-col += 1
                  if current-col >= cols-per-page {
                     current-col = 0
                     current-page += 1
                  }
               }
            }
         }
      }

      if current-row > 0 {
         current-row = 0
         if current-col + 1 < cols-per-page {
            current-col += 1
         } else {
            current-col = 0
            current-page += 1
         }
      }
   }

   (pages: current-page + 1, headings: headings)
}

#let template(doc, char-per-col: 20, cols-per-half-page: 10) = {
   let page-width = 29.7cm
    
    
  //  let raw-cols = calc.floor(page-width / cell-width)
  //  let odd-cols = if calc.rem(raw-cols, 2) == 0 { raw-cols - 1 } else { raw-cols }
   let cols-per-page = 2 * cols-per-half-page
   let total-cols = cols-per-page + 1
   let middle-col = calc.floor(total-cols / 2)

  let x-margin = 3cm
  let cell-width = (page-width - 2 * x-margin) / total-cols
    let cell-height = cell-width / 1.5
    let page-height = cell-height * char-per-col
    let font-size = cell-height
    let note-font-size = font-size * 0.5
   let punct-font-size = font-size * 0.45

  
  set page(paper: "a4", flipped: true, margin: (x: -x-margin, top: 4cm), fill: rgb(236,231,184,40%), header-ascent: 0%, footer-descent: 0pt)
  set text(font: "KingHwa_OldSong", size: 16pt)
  // set text(font: "STFangsong", size: 16pt)
  
   let content-blocks = extract-blocks(doc)
      .map(b => (..b, text: b.text.replace(regex("[ \n\t　]"), "")))
      .filter(b => b.text.len() > 0)


   let content-sim = simulate-layout(content-blocks, char-per-col, cols-per-page)
   let toc-items = content-sim.headings.filter(h => h.level <= 3)

   let build-toc-blocks = (page-offset: 0) => {
      let toc = ((kind: "heading", level: 1, text: "目錄"),)
      for item in toc-items {
         let indent = if item.level == 1 { 0 } else if item.level == 2 { 1 } else { 2 }
         let page-text = to-chinese-number(item.page + page-offset + 1)
         let toc-title = strip-heading-notes(item.title)
         toc.push((
            kind: "toc",
            level: 0,
            text: toc-title + "【" + page-text + "】",
            toc-indent: indent,
         ))
      }
      toc
   }

   let toc-pages = 0
   let toc-blocks = build-toc-blocks(page-offset: 0)
   let iter = 0
   while iter < 3 {
      toc-pages = simulate-layout(toc-blocks, char-per-col, cols-per-page).pages
      let next-toc = build-toc-blocks(page-offset: toc-pages)
      if next-toc == toc-blocks {
         break
      }
      toc-blocks = next-toc
      iter += 1
   }

   let blocks = toc-blocks + content-blocks

   let logical-col-to-physical = logical-col => if logical-col < middle-col { logical-col } else { logical-col + 1 }
   let col-x = logical-col => page-width - (logical-col-to-physical(logical-col) + 1) * cell-width
    
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

    // Helper to advance position
    let advance-pos = (n: 1) => {
       // We can't update captured vars directly in function, so return new state?
       // Typst variables are immutable. We must manage state in loop.
       // Refactoring: Use a state object or just update local vars in loop.
    }
    
    // We'll manage state manually in the loop.
    
   let seen-level1-heading = false
    for blk in blocks {
       if blk.kind == "heading" and blk.level == 1 {
          if seen-level1-heading {
             current-row = 0
             current-col = 0
             current-page += 1
          }
          let start-page = if current-col < cols-per-half-page { current-page } else { current-page + 1 }
          banxin-title-map.push((start: start-page, title: strip-heading-notes(blk.text), level: 1))
          seen-level1-heading = true
       }

       if blk.kind == "heading" and blk.level == 2 {
          if current-row > 0 {
             current-row = 0
             if current-col + 1 < cols-per-page {
                current-col += 1
             } else {
                current-col = 0
                current-page += 1
             }
          }
          if current-col + 1 < cols-per-page {
             current-col += 1
          } else {
             current-col = 0
             current-page += 1
          }
          let start-page = if current-col < cols-per-half-page { current-page } else { current-page + 1 }
          banxin-title-map.push((start: start-page, title: strip-heading-notes(blk.text), level: 2))
       }

       if blk.kind == "heading" and blk.level == 3 {
          if current-row > 0 {
             current-row = 0
             if current-col + 1 < cols-per-page {
                current-col += 1
             } else {
                current-col = 0
                current-page += 1
             }
          }
          if current-col + 1 < cols-per-page {
             current-col += 1
          } else {
             current-col = 0
             current-page += 1
          }
          if char-per-col > 2 {
             current-row = 2
          } else if char-per-col > 1 {
             current-row = 1
          }
       }

      if blk.at("toc-indent", default: 0) > 0 {
         let toc-indent = blk.at("toc-indent", default: 0)
         if toc-indent < char-per-col {
            current-row = toc-indent
         } else {
            current-row = char-per-col - 1
         }
      }

      let segments = parse-segments(blk.text)
       
       for segment in segments {
          if segment.type == "text" {
             for char in segment.content.clusters() {
                if is-punct(char) {
                   if has-last-anchor {
                      if pages-content.len() <= last-anchor-page {
                         while pages-content.len() <= last-anchor-page {
                            pages-content.push(())
                         }
                      }
                      pages-content.at(last-anchor-page).push(place(
                         top + left,
                         dx: last-anchor-x + last-anchor-w * 0.62,
                         dy: last-anchor-y + last-anchor-h * 0.62,
                         block(
                            width: last-anchor-w * 0.38,
                            height: last-anchor-h * 0.38,
                            align(center + horizon, text(size: punct-font-size, fill: red, char))
                         )
                      ))
                   }
                } else {
                   // Add regular character
                   let x-pos = col-x(current-col)
                   let y-pos = current-row * cell-height

                   if pages-content.len() <= current-page {
                      while pages-content.len() <= current-page {
                         pages-content.push(())
                      }
                   }

                   pages-content.at(current-page).push(place(
                      top + left, dx: x-pos, dy: y-pos,
                      block(width: cell-width, height: cell-height, align(center + horizon, text(size: font-size, char)))
                   ))

                   has-last-anchor = true
                   last-anchor-page = current-page
                   last-anchor-x = x-pos
                   last-anchor-y = y-pos
                   last-anchor-w = cell-width
                   last-anchor-h = cell-height

                   // Advance
                   current-row += 1
                   if current-row >= char-per-col {
                      current-row = 0; current-col += 1
                      if current-col >= cols-per-page { current-col = 0; current-page += 1 }
                   }
                }
             }

          } else { // Note
             let pre-note-has-anchor = has-last-anchor
             let pre-note-anchor-page = last-anchor-page
             let pre-note-anchor-x = last-anchor-x
             let pre-note-anchor-y = last-anchor-y
             let pre-note-anchor-w = last-anchor-w
             let pre-note-anchor-h = last-anchor-h

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
                if current-row >= char-per-col {
                   current-row = 0
                   current-col += 1
                   if current-col >= cols-per-page {
                      current-col = 0
                      current-page += 1
                   }
                }

                let rows-left = char-per-col - current-row
                let grids-remaining = grids-total - grids-processed
                let grids-in-segment = calc.min(rows-left, grids-remaining)

                let capacity = grids-in-segment * 4
                let chars-count = calc.min(total-chars - chars-processed, capacity)
                let segment-chars = note-chars.slice(chars-processed, chars-processed + chars-count)

                let right-count = if calc.rem(chars-count, 4) == 1 {
                   let k = 0
                   let remaining = chars-count
                   while remaining >= 4 {
                      remaining -= 4
                      k += 1
                   }
                   2 * k + 1
                } else if calc.rem(chars-count, 4) == 2 {
                   calc.floor(chars-count / 2)
                } else {
                   calc.min(chars-count, 2 * grids-in-segment)
                }

                if pages-content.len() <= current-page {
                   while pages-content.len() <= current-page {
                      pages-content.push(())
                   }
                }

                let base-x = col-x(current-col)
                let base-y = current-row * cell-height
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

                   pages-content.at(current-page).push(place(
                      top + left,
                      dx: char-x,
                      dy: char-y,
                      block(width: sub-w, height: sub-h, align(center + horizon, text(size: note-font-size, segment-chars.at(idx))))
                   ))

                   note-char-anchors.push((page: current-page, x: char-x, y: char-y, w: sub-w, h: sub-h))
                   has-last-anchor = true
                   last-anchor-page = current-page
                   last-anchor-x = char-x
                   last-anchor-y = char-y
                   last-anchor-w = sub-w
                   last-anchor-h = sub-h

                   idx += 1
                }

                chars-processed += chars-count
                grids-processed += grids-in-segment
                current-row += grids-in-segment

                if current-row >= char-per-col {
                   current-row = 0
                   current-col += 1
                   if current-col >= cols-per-page {
                      current-col = 0
                      current-page += 1
                   }
                }
             }

             for p in note-punct {
                let has-anchor = false
                let anchor-page = 0
                let anchor-x = 0pt
                let anchor-y = 0pt
                let anchor-w = cell-width
                let anchor-h = cell-height

                if p.target >= 0 and p.target < note-char-anchors.len() {
                   let a = note-char-anchors.at(p.target)
                   has-anchor = true
                   anchor-page = a.page
                   anchor-x = a.x
                   anchor-y = a.y
                   anchor-w = a.w
                   anchor-h = a.h
                } else if pre-note-has-anchor {
                   has-anchor = true
                   anchor-page = pre-note-anchor-page
                   anchor-x = pre-note-anchor-x
                   anchor-y = pre-note-anchor-y
                   anchor-w = pre-note-anchor-w
                   anchor-h = pre-note-anchor-h
                }

                if has-anchor {
                   if pages-content.len() <= anchor-page {
                      while pages-content.len() <= anchor-page {
                         pages-content.push(())
                      }
                   }
                   pages-content.at(anchor-page).push(place(
                      top + left,
                      dx: anchor-x + anchor-w * 0.62,
                      dy: anchor-y + anchor-h * 0.62,
                      block(
                         width: anchor-w * 0.38,
                         height: anchor-h * 0.38,
                         align(center + horizon, text(size: punct-font-size, fill: red, p.char))
                      )
                   ))
                }
             }
          }

       }
       
       if current-row > 0 {
          current-row = 0
          if current-col + 1 < cols-per-page {
             current-col += 1
          } else {
             current-col = 0
             current-page += 1
          }
       }
    }
    
    for (page-idx, content) in pages-content.enumerate() {
       let grid-lines = ()
       let frame-lines = ()

      let grid-left = page-width - total-cols * cell-width
       let grid-right = page-width
      let frame-offset = 5pt
       let frame-left = grid-left - frame-offset
       let frame-right = grid-right + frame-offset
       let v-shift = 2pt

       for col-idx in range(total-cols + 1) {
          let x-pos = page-width - (col-idx) * cell-width
          grid-lines.push(place(
             top + left, dx: x-pos, dy: 0pt,
             line(start: (0pt, -v-shift), end: (0pt, page-height + v-shift), stroke: 0.5pt )
          ))
       }

       frame-lines.push(place(
          top + left, dx: frame-left, dy: 0pt,
          polygon((0pt, -v-shift), (frame-right - frame-left, -v-shift), (frame-right - frame-left, page-height + v-shift), (0pt, page-height + v-shift), stroke: 2pt)
       ))

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
       )

          box(width: 100%, height: 100%, { grid-lines.join(); frame-lines.join(); content.join(); banxin })
          if page-idx < pages-content.len() - 1 { pagebreak(weak: true) }
    }
}
