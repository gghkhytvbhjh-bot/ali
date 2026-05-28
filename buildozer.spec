[app]
title = AidAdha Controle
package.name = aidadha
package.domain = org.aidadha
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 1.0
requirements = python3,kivy==2.2.1,kivymd==1.1.1,pillow,openpyxl
orientation = portrait
fullscreen = 1
android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.api = 33
android.minapi = 24
android.ndk = 25b
android.private_storage = True
android.archs = arm64-v8a
android.allow_backup = True

[buildozer]
log_level = 2
warn_on_root = 1
