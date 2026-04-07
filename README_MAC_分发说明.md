# Thesis Reviewer Mac 分发说明

## 先说结论

`build_thesis_reviewer_mac.command` 很小是正常的，因为它只是一个启动脚本，不是完整程序。

如果你要发给朋友，有两种方式：

## 方式一：发“源码打包包”

适合对方有 Mac，并且愿意自己点一次构建。

你需要发这几个文件：

- `thesis_reviewer_app.py`
- `thesis_reviewer_requirements.txt`
- `build_thesis_reviewer_mac.sh`
- `build_thesis_reviewer_mac.command`
- `README_MAC_分发说明.md`

对方在 Mac 上的操作步骤：

1. 先安装 Python 3。
2. 把以上文件放在同一个文件夹里。
3. 双击 `build_thesis_reviewer_mac.command`。
4. 如果系统拦截，就右键该文件，选择“打开”。
5. 等待构建完成。
6. 构建成功后，会生成：
   `dist_thesis_reviewer_mac/ThesisReviewerMac.app`
7. 双击 `ThesisReviewerMac.app` 即可运行。

如果双击 `.command` 没反应，也可以在终端中运行：

```bash
cd 到该文件夹
bash build_thesis_reviewer_mac.sh
```

## 方式二：发“已经构建好的 .app”

这是更适合普通朋友的方式。

流程是：

1. 找一台 Mac，先运行 `build_thesis_reviewer_mac.command` 或 `build_thesis_reviewer_mac.sh`
2. 生成 `ThesisReviewerMac.app`
3. 把整个 `ThesisReviewerMac.app` 压缩成 zip
4. 把这个 zip 发给朋友
5. 朋友解压后，双击 `.app` 运行

注意：

- `.app` 必须在 Mac 上构建，不能在 Windows 上直接生成真正可运行的 Mac 应用
- 如果 Mac 提示“无法打开”，通常是因为没有签名；可以右键应用，选择“打开”

## 方式三：用 GitHub Actions 自动构建

如果你自己没有 Mac，但想拿到可直接发给朋友的 `ThesisReviewerMac.app.zip`，这是最合适的方式。

你需要把这些文件上传到 GitHub 仓库：

- `thesis_reviewer_app.py`
- `thesis_reviewer_requirements.txt`
- `build_thesis_reviewer_mac.sh`
- `build_thesis_reviewer_mac.command`
- `.github/workflows/build-thesis-reviewer-mac.yml`

然后在 GitHub 上这样操作：

1. 打开仓库页面。
2. 点击 `Actions`。
3. 找到 `Build Thesis Reviewer Mac App`。
4. 点击 `Run workflow`。
5. 等待构建完成。
6. 进入该次运行记录。
7. 在 `Artifacts` 区域下载 `ThesisReviewerMac-app`。
8. 下载后得到 `ThesisReviewerMac.app.zip`。
9. 把这个 zip 直接发给朋友。

朋友收到后只需要：

1. 解压 zip
2. 得到 `ThesisReviewerMac.app`
3. 双击运行

如果系统拦截：

- 右键应用
- 选择“打开”

## 你现在最应该发什么

如果你朋友只是“帮你试运行”并且会一点电脑操作：

发一个包含以下文件的压缩包即可：

- `thesis_reviewer_app.py`
- `thesis_reviewer_requirements.txt`
- `build_thesis_reviewer_mac.sh`
- `build_thesis_reviewer_mac.command`
- `README_MAC_分发说明.md`

如果你朋友是最终使用者，不想折腾环境：

不要发源码包。
应该先在一台 Mac 上构建出 `ThesisReviewerMac.app`，或者用 GitHub Actions 自动构建出 `ThesisReviewerMac.app.zip`，然后把成品发给他。

## 当前版本的 Mac 限制

Mac 版已经支持：

- 上传论文
- 调用 API 评审
- 生成详细批注文档
- 生成评分与 300 字总结文档

但目前仍不支持：

- 像 Windows 版那样直接生成“原文内嵌 Word 批注版”

这是因为该功能依赖 Windows 下的 Microsoft Word 自动化。
