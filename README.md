# 人教版高中物理课件：热力学第一定律

这是一个 16:9 的高中物理网页课件项目，主题为“热力学第一定律”。课件使用原生 HTML、CSS 和 JavaScript 制作，可直接用浏览器全屏投屏。

## 文件结构

```text
.
├── index.html
├── README.md
├── AGENTS.md
└── slides/
    ├── slide-01.html
    ├── slide-02.html
    ├── slide-03.html
    ├── slide-04.html
    ├── slide-05.html
    ├── slide-06.html
    ├── slide-07.html
    └── slide-08.html
```

## 运行与预览

无需安装依赖。推荐使用本地静态服务器预览：

```bash
python3 -m http.server 8000
```

然后打开：

```text
http://localhost:8000/
```

也可以直接用浏览器打开 `index.html`，每个 `slides/slide-XX.html` 页面也支持独立打开。

## 操作方式

- 点击当前页面：逐步出现课堂讲解内容；
- `←` / `→`：上一页 / 下一页；
- `F`：进入或退出全屏；
- `R`：重置当前页动画；
- 首页和每个独立课件页底部控件均支持上一页、下一页、页码跳转、重置和全屏。

## 物理约定

全套课件统一使用：

```text
ΔU = Q + W
```

- `Q > 0`：系统吸热；
- `Q < 0`：系统放热；
- `W > 0`：外界对系统做功；
- `W < 0`：系统对外界做功。

## 技术说明

- 原生 HTML、CSS、JavaScript；
- 不使用 React、Vue、Vite 等框架；
- 不安装任何依赖；
- 不调用外部图片、字体、CSS 或 JavaScript；
- 图形使用内联 SVG；
- 不使用 Canvas；
- 固定 16:9 比例，适配 1920×1080 和 1366×768；
- 页面不应出现横向或纵向滚动条。
