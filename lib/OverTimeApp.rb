#coding: utf-8

# ======================================
# 時間外申請書作成スクリプト
# ======================================

require "win32ole"
require "date"
require "fileutils"
require "json"
require "tmpdir"
require "./lib/xlsbook.rb"
require "./lib/config.rb"

# ======================================
# 時間外申請書作成クラス
# ======================================
class OverTimeApp

  def initialize

    # 設定ファイルロード
    @config = Config.load("./config/config")

    # 設定ファイルの検証
    validation_config

  end

  # ======================================
  # 値ファイルロード
  # ======================================
  def load_app_file

    JSON.load(
      open(@config.app_file, "r", :encoding => "UTF-8")
    )

  end

  # ======================================
  # 設定ファイル検証
  # ======================================
  def validation_config()

    unless(File.exist?(@config.app_file))
      raise "データファイルが見つかりません"
    end

    if(@config.template_file.nil? || @config.template_file.empty?)
      raise "設定ファイルにテンプレートファイルの定義が見つかりません"
    end

    unless(File.exist?(@config.template_file))
      raise "テンプレートファイルが見つかりません"
    end

    if(@config.sheet_name.nil? || @config.sheet_name.empty?)
      raise "設定ファイルにシート名の定義が見つかりません"
    end

    if(@config.output_xls_file.nil? || @config.output_xls_file.empty?)
      raise "設定ファイルにExcelファイル名の定義が見つかりません"
    end

    if(@config.app_file.nil? || @config.app_file.empty?)
      raise "設定ファイルにデータファイルの定義が見つかりません"
    end


  end

  # --------------------------------------
  # 申請ファイル作成
  # --------------------------------------
  def create_xls(xls_file)

    values = load_app_file()

    xls = XlsBook.new

    sheet = @config.sheet_name

    xls.load_book(xls_file, sheet) do |sh|

      ymd = Date.today

      # 申請日
      sh.cells(6,19).value = ymd.year
      sh.cells(6,23).value = ymd.mon
      sh.cells(6,26).value = ymd.mday

      # 残業累積
      sh.cells(7, 25).value = values["ruiseki"]

      # 年月日時分
      sh.cells(19, 6).value = ymd.year
      sh.cells(19, 12).value = ymd.mon
      sh.cells(19, 18).value = ymd.mday
      sh.cells(22, 17).value = values["end_hour"]
      sh.cells(22, 21).value = values["end_minite"]

      # 業務内容 残業作業
      sh.cells(30, 6).value = values["task_name"]

      # 業務内容 納期
      sh.cells(30, 18).value = values["nouki"]

      # 業務内容 工数
      sh.cells(30, 22).value = values["kousu"]

      # 業務内容 進捗
      sh.cells(30, 25).value = values["progress"]

      # 申請理由
      sh.cells(36, 6).value = values["reason"].encode("UTF-8")

    end

  end

  # --------------------------------------
  # 申請ファイル作成
  # --------------------------------------
  def generate()

    # テンプレートファイルから一時ファイルを作成する
    s = rand(9999999999).to_s.rjust(10, "0")
    temp_file = "#{Dir.tmpdir}/#{s}.xls"
    FileUtils.cp @config.template_file, temp_file

    # 申請書ファイル(一時)作成
    create_xls temp_file

    #ここでWaitさせないとエラーになる
    sleep 5

    # 申請書ファイル(一時)から申請書ファイル(正)を生成
    file_name = DateTime.now.strftime(@config.output_xls_file.encode("windows-31j"))
    FileUtils.mv temp_file, "#{@config.output_dir}/#{file_name}"

  end

end

# MAIN-ROUTINE
overtimeapp = OverTimeApp.new

overtimeapp.generate
