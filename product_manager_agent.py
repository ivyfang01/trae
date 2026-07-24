#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
产品经理智能体 v1.0
具备10年产品经验的AI助手，擅长：
- 用户洞察分析
- PRD文档撰写
- 数据分析与决策
"""

import json
from typing import Dict, List, Any
from datetime import datetime

class ProductManagerAgent:
    """
    产品经理智能体核心类
    模拟10年经验产品经理的思考方式和工作流程
    """
    
    def __init__(self):
        self.agent_info = {
            "name": "PM智能助手",
            "experience": "10年",
            "specialties": ["用户洞察", "PRD撰写", "数据分析", "需求挖掘"],
            "tools": ["用户访谈分析", "竞品分析", "数据看板", "原型设计"]
        }
        
    def analyze_user_needs(self, user_input: str) -> Dict[str, Any]:
        """
        用户需求分析 - 深度挖掘用户真实需求
        """
        analysis = {
            "需求类型": self._classify_requirement(user_input),
            "核心诉求": self._extract_core_request(user_input),
            "潜在需求": self._identify_latent_needs(user_input),
            "优先级评估": self._evaluate_priority(user_input),
            "风险提示": self._identify_risks(user_input)
        }
        return analysis
    
    def write_prd(self, requirements: Dict[str, Any]) -> str:
        """
        PRD文档撰写 - 生成结构化的产品需求文档
        """
        prd_template = f"""
# {requirements.get('产品名称', '未命名功能')} - PRD文档

---

## 1. 需求背景

### 1.1 业务背景
{requirements.get('业务背景', '待补充')}

### 1.2 用户场景
{requirements.get('用户场景', '待补充')}

---

## 2. 需求概述

| 属性 | 说明 |
|------|------|
| 需求ID | {requirements.get('需求ID', 'PRD-001')} |
| 需求名称 | {requirements.get('需求名称', '')} |
| 所属模块 | {requirements.get('所属模块', '未指定')} |
| 优先级 | {requirements.get('优先级', '中')} |
| 需求来源 | {requirements.get('需求来源', '业务提需')} |

---

## 3. 功能需求

### 3.1 核心功能
{self._format_features(requirements.get('功能列表', []))}

### 3.2 交互逻辑
{requirements.get('交互逻辑', '待补充')}

---

## 4. 页面原型描述

{requirements.get('原型描述', '待补充')}

---

## 5. 数据字段定义

{self._format_data_fields(requirements.get('数据字段', []))}

---

## 6. 非功能需求

- **性能要求**：{requirements.get('性能要求', '页面加载≤2s')}
- **兼容性**：{requirements.get('兼容性', '主流浏览器兼容')}
- **安全性**：{requirements.get('安全性', '数据传输加密')}

---

## 7. 验收标准

{self._format_acceptance_criteria(requirements.get('验收标准', []))}

---

**文档版本**: V1.0  
**创建时间**: {datetime.now().strftime('%Y-%m-%d')}
"""
        return prd_template.strip()
    
    def analyze_data(self, metrics: Dict[str, float], context: str = "") -> Dict[str, Any]:
        """
        数据分析 - 基于指标提供洞察和建议
        """
        insights = {
            "概览": self._summarize_metrics(metrics),
            "趋势分析": self._analyze_trends(metrics),
            "问题识别": self._identify_problems(metrics),
            "优化建议": self._generate_recommendations(metrics, context),
            "预期收益": self._estimate_impact(metrics)
        }
        return insights
    
    def conduct_user_research(self, interview_notes: List[str]) -> Dict[str, Any]:
        """
        用户研究分析 - 从访谈记录中提取洞察
        """
        research_result = {
            "用户画像": self._build_user_profile(interview_notes),
            "痛点汇总": self._extract_pain_points(interview_notes),
            "需求排序": self._prioritize_needs(interview_notes),
            "机会点": self._identify_opportunities(interview_notes)
        }
        return research_result
    
    # ============ 私有方法 - 核心分析逻辑 ============
    
    def _classify_requirement(self, input_text: str) -> str:
        """分类需求类型"""
        keywords = {
            "表单": "功能开发",
            "页面": "界面设计",
            "报告": "数据分析",
            "需求": "需求分析",
            "原型": "产品设计",
            "bug": "缺陷修复",
            "性能": "性能优化",
            "安全": "安全需求"
        }
        for kw, category in keywords.items():
            if kw in input_text:
                return category
        return "其他需求"
    
    def _extract_core_request(self, input_text: str) -> str:
        """提取核心诉求"""
        return f"用户需要{input_text.strip()}"
    
    def _identify_latent_needs(self, input_text: str) -> List[str]:
        """识别潜在需求"""
        latent_map = {
            "表单": ["数据收集完整性", "用户体验流畅性", "数据安全性"],
            "页面": ["响应式适配", "加载性能", "视觉一致性"],
            "报告": ["数据准确性", "可视化呈现", "导出功能"]
        }
        for kw, needs in latent_map.items():
            if kw in input_text:
                return needs
        return ["系统稳定性", "扩展性需求"]
    
    def _evaluate_priority(self, input_text: str) -> str:
        """评估优先级"""
        high_keywords = ["紧急", "必须", "立刻", "马上"]
        low_keywords = ["优化", "改进", "建议"]
        
        if any(kw in input_text for kw in high_keywords):
            return "高"
        elif any(kw in input_text for kw in low_keywords):
            return "低"
        return "中"
    
    def _identify_risks(self, input_text: str) -> List[str]:
        """识别风险"""
        risks = []
        if "手机号" in input_text:
            risks.append("隐私合规风险：需确保手机号加密存储")
        if "提交" in input_text:
            risks.append("数据重复提交风险：需增加防重复机制")
        return risks if risks else ["暂无明显风险"]
    
    def _format_features(self, features: List[str]) -> str:
        """格式化功能列表"""
        if not features:
            return "待补充"
        return "\n".join(f"- {i+1}. {feature}" for i, feature in enumerate(features))
    
    def _format_data_fields(self, fields: List[Dict]) -> str:
        """格式化数据字段"""
        if not fields:
            return "待补充"
        table = "| 字段名 | 类型 | 必填 | 说明 |\n|--------|------|------|------|\n"
        for field in fields:
            table += f"| {field.get('name', '')} | {field.get('type', '')} | {field.get('required', '')} | {field.get('desc', '')} |\n"
        return table
    
    def _format_acceptance_criteria(self, criteria: List[str]) -> str:
        """格式化验收标准"""
        if not criteria:
            return "待补充"
        return "\n".join(f"{i+1}. {criterion}" for i, criterion in enumerate(criteria))
    
    def _summarize_metrics(self, metrics: Dict) -> str:
        """汇总指标"""
        total = sum(metrics.values())
        avg = total / len(metrics) if metrics else 0
        return f"共{len(metrics)}项指标，总和{total:.2f}，平均值{avg:.2f}"
    
    def _analyze_trends(self, metrics: Dict) -> List[str]:
        """趋势分析"""
        trends = []
        for name, value in metrics.items():
            if value > 80:
                trends.append(f"{name}表现优秀（{value}分）")
            elif value < 60:
                trends.append(f"{name}存在短板（{value}分）")
        return trends if trends else ["数据平稳，无明显趋势"]
    
    def _identify_problems(self, metrics: Dict) -> List[str]:
        """问题识别"""
        problems = []
        threshold = 60
        for name, value in metrics.items():
            if value < threshold:
                problems.append(f"{name}低于阈值{threshold}，当前{value}分")
        return problems if problems else ["未发现明显问题"]
    
    def _generate_recommendations(self, metrics: Dict, context: str) -> List[str]:
        """生成建议"""
        recommendations = []
        if "转化率" in metrics and metrics["转化率"] < 70:
            recommendations.append("建议优化表单填写流程，减少用户流失")
        if "满意度" in metrics and metrics["满意度"] < 80:
            recommendations.append("建议收集用户反馈，针对性改进体验")
        return recommendations if recommendations else ["当前数据良好，建议保持"]
    
    def _estimate_impact(self, metrics: Dict) -> str:
        """预估收益"""
        improvement = sum(1 for v in metrics.values() if v < 70) * 15
        return f"预计优化后可提升指标约{improvement}%，带来正向业务影响"
    
    def _build_user_profile(self, notes: List[str]) -> Dict[str, Any]:
        """构建用户画像"""
        profile = {
            "主要人群": "业务用户",
            "核心目标": "高效完成任务",
            "使用场景": "办公环境"
        }
        return profile
    
    def _extract_pain_points(self, notes: List[str]) -> List[str]:
        """提取痛点"""
        pain_points = []
        for note in notes:
            if "麻烦" in note or "复杂" in note:
                pain_points.append("操作流程复杂")
            if "慢" in note or "卡顿" in note:
                pain_points.append("系统响应慢")
        return pain_points if pain_points else ["未明确提及痛点"]
    
    def _prioritize_needs(self, notes: List[str]) -> List[str]:
        """需求排序"""
        return ["核心功能完善", "体验优化", "性能提升"]
    
    def _identify_opportunities(self, notes: List[str]) -> List[str]:
        """识别机会点"""
        return ["流程优化", "智能化升级", "数据驱动决策"]

# ============ 使用示例 ============

def demo_agent():
    """产品经理智能体使用演示"""
    agent = ProductManagerAgent()
    
    print("="*60)
    print("🎯 产品经理智能体 v1.0")
    print("="*60)
    
    # 1. 用户需求分析演示
    print("\n📊 【演示1】用户需求分析")
    user_input = "帮我做一个表单页面，需要手机号和品牌字段"
    analysis = agent.analyze_user_needs(user_input)
    print(json.dumps(analysis, ensure_ascii=False, indent=2))
    
    # 2. PRD撰写演示
    print("\n📝 【演示2】PRD文档生成")
    requirements = {
        "产品名称": "完善信息表单",
        "业务背景": "为业务员提供信息完善入口，确保数据完整性",
        "用户场景": "业务员首次登录或信息变更时填写",
        "需求ID": "PRD-FORM-001",
        "需求名称": "完善信息表单页面",
        "所属模块": "用户中心",
        "优先级": "高",
        "功能列表": [
            "手机号输入与验证",
            "品牌下拉选择",
            "表单提交与重置",
            "数据格式校验"
        ],
        "交互逻辑": "填写→验证→提交→成功提示",
        "数据字段": [
            {"name": "phone", "type": "String(11)", "required": "是", "desc": "业务员手机号"},
            {"name": "brand", "type": "Enum", "required": "是", "desc": "品牌归属"}
        ],
        "验收标准": [
            "手机号格式正确才能提交",
            "品牌为必选项",
            "取消按钮可重置表单",
            "页面支持响应式"
        ]
    }
    prd = agent.write_prd(requirements)
    print(prd[:1000] + "..." if len(prd) > 1000 else prd)
    
    # 3. 数据分析演示
    print("\n📈 【演示3】数据分析")
    metrics = {
        "转化率": 65.5,
        "满意度": 78.3,
        "留存率": 85.2,
        "响应速度": 92.1
    }
    data_analysis = agent.analyze_data(metrics, context="表单页面")
    print(json.dumps(data_analysis, ensure_ascii=False, indent=2))
    
    # 4. 用户研究演示
    print("\n🔍 【演示4】用户研究分析")
    interview_notes = [
        "用户觉得填写表单太麻烦",
        "希望能快速完成",
        "系统有时候响应慢"
    ]
    research = agent.conduct_user_research(interview_notes)
    print(json.dumps(research, ensure_ascii=False, indent=2))
    
    print("\n" + "="*60)
    print("✅ 演示完成！")
    print("="*60)

if __name__ == "__main__":
    demo_agent()