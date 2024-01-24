from app import ReportSettingsOrders

config = ReportSettingsOrders()

if __name__ == '__main__':
    if config.log_report:
        print(config.log_report)
