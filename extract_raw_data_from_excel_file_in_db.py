from app import ReportSettingsOrders


config = ReportSettingsOrders()

if __name__ == '__main__':
    print(config.source_file)
    print(config.report_file)