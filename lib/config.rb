#coding: utf-8

# =====================================
# 設定ファイル読み込みクラス
# =====================================
class Config
  def self.load(config_file)
    c = new()
    c.instance_eval File.read(config_file, encoding: 'UTF-8')
    c
  end
  attr_reader :home_dir
  attr_reader :app_file
  attr_reader :template_file
  attr_reader :sheet_name
  attr_reader :output_dir
  attr_reader :output_xls_file
end
