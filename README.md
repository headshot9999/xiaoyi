# react-native-view-shot Android 构建配置修复

## 问题描述

在使用较旧版本的 Android Gradle Plugin (AGP < 7.0) 构建项目时，`react-native-view-shot` 的 `android/build.gradle` 会因以下错误而构建失败：

```
A problem occurred evaluating project ':react-native-view-shot'.
> Could not get unknown property 'com' for extension 'android' of type com.android.build.gradle.LibraryExtension.
```

### 根本原因

`build.gradle` 中第 14 行引用了 `com.android.Version.ANDROID_GRADLE_PLUGIN_VERSION`，该类仅在 AGP 7.0+ 中可用。当使用 AGP < 7.0（配合 Gradle 6.9）时，此引用会导致构建失败。

## 修复方案

将 `build.gradle` 中对 `com.android.Version.ANDROID_GRADLE_PLUGIN_VERSION` 的直接引用替换为一个安全的版本检测方法，该方法兼容新旧两个版本的 AGP：

1. 使用 `try-catch` 块安全地获取 AGP 版本号
2. 对于 AGP < 7.0（该类不存在），自动跳过 `namespace` 设置（因为旧版本不需要）
3. 保持对 AGP >= 7.0 的完全兼容

## 修改的文件

- `android/build.gradle` — 补丁文件，展示了推荐的修改方式

## 如何应用修复

将 `android/build.gradle` 中的内容应用到你项目中 `node_modules/react-native-view-shot/android/build.gradle` 文件。

### 方式一：使用 patch-package（推荐）

1. 手动编辑 `node_modules/react-native-view-shot/android/build.gradle`，按照本仓库 `android/build.gradle` 中的修改进行更新
2. 运行 `npx patch-package react-native-view-shot`
3. 将生成的 patch 文件提交到版本控制

### 方式二：升级依赖

升级 Android Gradle Plugin 到 7.0+ 和 Gradle 到 7.0+，这样就不需要修补了。
