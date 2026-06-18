# Medical Writing Timeline

这是一个纯静态网页版本，可以直接部署到 GitHub Pages、Netlify 或 Vercel。

## GitHub Pages 发布

1. 在 GitHub 新建一个仓库，例如 `medical-writing-timeline`。
2. 把本文件夹里的所有文件上传到仓库根目录：
   - `index.html`
   - `styles.css`
   - `app.js`
   - `timeline-core.js`
   - `holiday-data.js`
   - `.nojekyll`
3. 打开仓库 `Settings` -> `Pages`。
4. `Build and deployment` 选择：
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/ (root)`
5. 保存后等待 1-2 分钟，GitHub 会生成固定网址。

## 数据说明

项目数据保存在当前浏览器本地。换电脑使用时：

- 在旧电脑点击 `导出项目库`
- 在新电脑打开固定网址后点击 `导入项目库`

