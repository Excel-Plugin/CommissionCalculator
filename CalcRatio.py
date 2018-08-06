import InterfaceModule

class CalcRatio:
    def __init__(self):
        self.header=["开始时间","结束时间","规则名","业务员","固定比例","0-60"
            , "61-120","121-150","151-180","结束时间","切削液","切削油","结束时间"
            , "其他","售后占比","业务员2","固定比例2"]
        im=InterfaceModule.Easyexcel()
        self.ruleTitle, self.rules=im.get_sheet("规则")
        pass