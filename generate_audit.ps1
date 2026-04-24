
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection

function AddHeading($text, $level) {
    $sel.Style = $doc.Styles("Heading $level")
    $sel.TypeText($text)
    $sel.TypeParagraph()
}

function AddBody($text) {
    $sel.Style = $doc.Styles("Normal")
    $sel.TypeText($text)
    $sel.TypeParagraph()
}

function AddBullet($text) {
    $sel.Style = $doc.Styles("List Bullet")
    $sel.TypeText($text)
    $sel.TypeParagraph()
}

function AddPageBreak() {
    $sel.InsertBreak(7)
}

# COVER PAGE
$sel.Style = $doc.Styles("Title")
$sel.TypeText("INTERNAL AUDIT REPORT")
$sel.TypeParagraph()
$sel.Style = $doc.Styles("Subtitle")
$sel.TypeText("Governance and Oversight Review")
$sel.TypeParagraph()
AddBody ""
AddBody "Regulatory Framework: Central Bank of the UAE (CBUAE)"
AddBody "Applicable Business Lines: Buy Now Pay Later (BNPL) and Stored Value Facilities (SVF)"
AddBody "Prepared By: Internal Audit Function"
AddBody ("Date: " + (Get-Date -Format "dd MMMM yyyy"))
AddBody "Classification: Confidential - For Internal Use Only"
AddPageBreak

# SECTION 1
AddHeading "1. PURPOSE AND SCOPE" 1
AddBody "This internal audit covers the Governance and Oversight framework of a UAE-based fintech company licensed by the Central Bank of the UAE (CBUAE) to operate Buy Now Pay Later (BNPL) and Stored Value Facility (SVF) products. The audit assesses compliance with applicable CBUAE regulations, standards, and circulars governing corporate governance, risk management, consumer protection, AML/CFT, and oversight obligations."
AddBody ""
AddBody "Audit Period: [Insert Audit Period]"
AddBody "Audit Universe: Board, Senior Management, Compliance, Risk, Operations, Finance"
AddBody "Risk Rating: High - Regulated Financial Services"
AddPageBreak

# SECTION 2
AddHeading "2. APPLICABLE CBUAE LAWS AND REGULATIONS" 1

AddHeading "2.1 Primary Legislation" 2
AddBullet "UAE Federal Law No. 14 of 2018 - Regarding the Central Bank and Organization of Financial Institutions and Activities"
AddBullet "UAE Federal Decree-Law No. 20 of 2018 on Anti-Money Laundering and Combating the Financing of Terrorism"
AddBullet "UAE Federal Law No. 2 of 2019 - Consumer Protection Law"
AddBullet "UAE Federal Decree-Law No. 45 of 2021 - Personal Data Protection Law (PDPL)"

AddHeading "2.2 CBUAE Regulations and Standards" 2
AddBullet "CBUAE Stored Value Facilities (SVF) Regulation - November 2020 (Circular No. CBUAE/BSD/N/2020/4966)"
AddBullet "CBUAE Retail Payment Services and Card Schemes (RPSCS) Regulation - July 2021 (Circular No. 15/2021)"
AddBullet "CBUAE Finance Companies Regulation - December 2023 (includes BNPL as Restricted License Finance Companies)"
AddBullet "CBUAE Consumer Protection Regulation - 2021"
AddBullet "CBUAE Corporate Governance Standards for Licensed Financial Institutions - 2023"
AddBullet "CBUAE Operational Risk Management Guidance - 2022"
AddBullet "CBUAE Cyber Risk Management Guidance - 2021"
AddBullet "CBUAE AML/CFT Guidelines for Licensed Financial Institutions"
AddBullet "CBUAE Outsourcing Regulation - 2021"
AddBullet "CBUAE Business Continuity Management (BCM) Standards"

AddHeading "2.3 Supporting International Frameworks" 2
AddBullet "Cabinet Resolution No. 10 of 2019 - Executive Regulation on AML/CFT"
AddBullet "FATF 40 Recommendations as adopted by the UAE"
AddBullet "Basel Committee on Banking Supervision (BCBS) Principles for Sound Corporate Governance"

AddPageBreak

# SECTION 3
AddHeading "3. KEY APPLICABLE ARTICLES" 1

AddHeading "3.1 SVF Regulation - Key Articles" 2
AddBullet "Article 3 - Licensing: All SVF issuers must obtain a licence from CBUAE prior to offering SVF services."
AddBullet "Article 5 - Capital Requirements: Minimum paid-up capital of AED 50 million; maintain minimum liquid assets of 10% of aggregate outstanding float."
AddBullet "Article 7 - Safeguarding of Funds: SVF float must be held in a designated safeguarding account at a licensed UAE bank, separate from operational funds."
AddBullet "Article 8 - Float Reconciliation: Daily reconciliation of float balances against SVF accounts required."
AddBullet "Article 10 - Governance: Board oversight required; defined roles for CEO, Compliance Officer, Risk Officer, and MLRO."
AddBullet "Article 12 - Risk Management: Comprehensive risk management framework covering operational, credit, liquidity, and reputational risks."
AddBullet "Article 14 - AML/CFT Compliance: Full compliance with Federal Decree-Law No. 20 of 2018; dedicated MLRO mandatory."
AddBullet "Article 15 - Consumer Protection: Clear disclosure of fees, terms, and conditions; robust complaint-handling mechanism."
AddBullet "Article 17 - Cybersecurity: Compliance with CBUAE Cyber Risk Management Guidance; UAE data residency mandatory."
AddBullet "Article 19 - Outsourcing: Board-approved outsourcing policy; prior CBUAE notification for material outsourcing."
AddBullet "Article 21 - Reporting Obligations: Quarterly regulatory reporting to CBUAE; immediate notification of material events."
AddBullet "Article 23 - CBUAE Supervisory Powers: CBUAE may conduct on-site inspections and impose corrective actions."

AddHeading "3.2 BNPL / Finance Companies Regulation - Key Articles" 2
AddBullet "Article 2 - Scope: BNPL providers classified as Restricted License Finance Companies unless operating as bank agents."
AddBullet "Article 4 - Minimum Capital: AED 20 million or 5% of outstanding lending volume (whichever is higher)."
AddBullet "Article 6 - Credit Limit Cap: Maximum credit exposure per borrower capped at AED 20,000 or 3 months verified net income."
AddBullet "Article 8 - Credit Assessment: Credit bureau check mandatory for exposures exceeding AED 5,000."
AddBullet "Article 9 - Fee Cap: Total fees and charges, including late payment fees, not to exceed 30% of the initial loan amount."
AddBullet "Article 11 - Disclosure Requirements: Full pre-contractual disclosure of all terms, fees, and repayment schedules."
AddBullet "Article 13 - Debt Collection: Fair debt collection practices compliant with Consumer Protection Regulation."
AddBullet "Article 15 - Governance: Board-approved credit policy; independent risk and compliance functions."
AddBullet "Article 17 - AML/CFT: Customer due diligence, transaction monitoring, and STR reporting to UAE FIU (goAML)."

AddHeading "3.3 RPSCS Regulation - Key Articles" 2
AddBullet "Article 4 - Licensing: Payment service providers must hold a CBUAE payment service licence (Category I-IV)."
AddBullet "Article 8 - Capital Requirements: Minimum capital based on licence category (AED 1M to AED 50M)."
AddBullet "Article 12 - Governance: Board-approved governance framework; fit and proper requirements for board and senior management."
AddBullet "Article 14 - Risk Management: Enterprise risk management framework; independent risk oversight function."
AddBullet "Article 16 - Consumer Protection: Transparent pricing; dispute resolution; 5-day complaint resolution SLA."
AddBullet "Article 18 - AML/CFT: Risk-based AML/CFT programme; sanctions screening against UN, OFAC, and UAE local lists."
AddBullet "Article 20 - Cybersecurity: Cyber risk management framework; penetration testing annually."
AddBullet "Article 22 - Outsourcing: Board-approved outsourcing policy; due diligence on third-party service providers."
AddBullet "Article 25 - Reporting: Monthly, quarterly, and annual regulatory returns to CBUAE."

AddHeading "3.4 Corporate Governance Standards - Key Articles" 2
AddBullet "Article 3 - Board Composition: Minimum independent directors; no executive chairman; separation of board and management."
AddBullet "Article 5 - Board Responsibilities: Board sets risk appetite, strategic direction, and oversees internal controls."
AddBullet "Article 7 - Board Committees: Mandatory committees include Audit, Risk, and Compliance."
AddBullet "Article 9 - Fit and Proper: All board members and senior management meet CBUAE fit and proper criteria."
AddBullet "Article 11 - Conflict of Interest: Written conflict of interest policy; register maintained and reviewed by board."
AddBullet "Article 13 - Internal Audit: Independent internal audit function reporting directly to Audit Committee."
AddBullet "Article 15 - Compliance Function: Chief Compliance Officer appointed; annual compliance plan approved by board."
AddBullet "Article 17 - Risk Function: Chief Risk Officer appointed; Risk Management Framework reviewed annually."
AddBullet "Article 19 - Remuneration: Remuneration policy aligned to risk; no incentivisation of excessive risk-taking."

AddPageBreak

# SECTION 4
AddHeading "4. INTERNAL AUDIT CHECKLIST - GOVERNANCE AND OVERSIGHT" 1
AddBody "Instructions: Mark each item as: C = Compliant | PC = Partially Compliant | NC = Non-Compliant | NA = Not Applicable"
AddBody ""

# 4.1
AddHeading "4.1 Board and Corporate Governance" 2
AddBullet "1. Is there a Board-approved Corporate Governance Policy reviewed at least annually? | Ref: Corp.Gov Art.5 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "2. Does Board composition meet CBUAE requirements (independent directors, no executive chairman)? | Ref: Corp.Gov Art.3 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "3. Are Board members and Senior Management assessed against CBUAE Fit and Proper criteria? | Ref: Corp.Gov Art.9 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "4. Is there clear documented separation between Board oversight and executive management? | Ref: Corp.Gov Art.5 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "5. Are mandatory Board committees established (Audit, Risk, Compliance)? | Ref: Corp.Gov Art.7 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "6. Do Board committee charters define roles, responsibilities, meeting frequency, and quorum? | Ref: Corp.Gov Art.7 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "7. Are Board and committee meeting minutes maintained and approved in a timely manner? | Ref: Corp.Gov Art.5 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "8. Is there a Board-approved Risk Appetite Statement (RAS) reviewed annually? | Ref: Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "9. Is the Board provided with regular MI packs covering key risk and performance indicators? | Ref: Corp.Gov Art.5 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "10. Is there a Conflict of Interest Policy and a register of disclosed conflicts maintained? | Ref: Corp.Gov Art.11 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.2
AddHeading "4.2 Licensing and Regulatory Standing" 2
AddBullet "11. Does the company hold a valid CBUAE licence for BNPL and SVF operations? | Ref: SVF Art.3; BNPL Art.2 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "12. Are all licensed products and services within the scope of the CBUAE licence? | Ref: SVF Art.3; BNPL Art.2 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "13. Are licence conditions and regulatory undertakings tracked for compliance? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "14. Are material changes (products, ownership, key personnel) notified to CBUAE within required timeframes? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "15. Is a regulatory correspondence register maintained and reviewed by Compliance? | Ref: SVF Art.23 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.3
AddHeading "4.3 Capital and Financial Adequacy" 2
AddBullet "16. Does the company maintain minimum paid-up capital (AED 50M for SVF; AED 20M for BNPL)? | Ref: SVF Art.5; BNPL Art.4 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "17. Is capital adequacy monitored on an ongoing basis with Board-level reporting? | Ref: SVF Art.5; Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "18. Are liquid assets maintained at minimum 10% of aggregate SVF float outstanding? | Ref: SVF Art.5 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "19. Is there a Capital Adequacy Policy and contingency plan for capital shortfalls? | Ref: SVF Art.5; BNPL Art.4 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.4
AddHeading "4.4 SVF Float Safeguarding and Reconciliation" 2
AddBullet "20. Are SVF customer funds held in a designated safeguarding account at a licensed UAE bank? | Ref: SVF Art.7 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "21. Is the safeguarding account segregated from operational funds? | Ref: SVF Art.7 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "22. Is daily reconciliation of float balances against SVF customer accounts performed and evidenced? | Ref: SVF Art.8 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "23. Are reconciliation exceptions escalated and resolved within defined SLAs? | Ref: SVF Art.8 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "24. Is there a Board-approved Safeguarding Policy reviewed at least annually? | Ref: SVF Art.7 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.5
AddHeading "4.5 BNPL Credit Risk and Consumer Limits" 2
AddBullet "25. Is a Board-approved Credit Policy in place covering underwriting, limits, and collections? | Ref: BNPL Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "26. Are BNPL credit limits capped at AED 20,000 or 3 months net income (whichever is lower)? | Ref: BNPL Art.6 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "27. Is a credit bureau check performed for all exposures exceeding AED 5,000? | Ref: BNPL Art.8 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "28. Are total fees and charges capped at 30% of the initial loan amount? | Ref: BNPL Art.9 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "29. Is income verification performed and documented before granting credit? | Ref: BNPL Art.6 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "30. Is the credit decisioning process documented, auditable, and tested? | Ref: BNPL Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "31. Are NPL trends monitored and reported to the Board and Risk Committee? | Ref: BNPL Art.15 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.6
AddHeading "4.6 Compliance Function" 2
AddBullet "32. Is a Chief Compliance Officer (CCO) appointed and independent from business lines? | Ref: Corp.Gov Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "33. Is an annual Compliance Monitoring Plan approved by Board or Audit Committee? | Ref: Corp.Gov Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "34. Are compliance monitoring results reported to the Board at least quarterly? | Ref: Corp.Gov Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "35. Is a Regulatory Horizon Scanning process in place to track upcoming CBUAE changes? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "36. Are compliance breaches logged, investigated, and remediated with root-cause analysis? | Ref: Corp.Gov Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "37. Is there a documented Policies and Procedures review cycle of at least annually? | Ref: Corp.Gov Art.15 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.7
AddHeading "4.7 Risk Management Framework" 2
AddBullet "38. Is a Chief Risk Officer (CRO) appointed and independent from business lines? | Ref: Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "39. Is a Board-approved Risk Management Framework in place and reviewed annually? | Ref: Corp.Gov Art.17; SVF Art.12 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "40. Does the RMF cover operational, credit, liquidity, market, reputational, and technology risks? | Ref: SVF Art.12 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "41. Is a Risk Register maintained with risks rated, owned, and mitigated? | Ref: SVF Art.12 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "42. Are Key Risk Indicators (KRIs) defined, monitored, and reported to the Risk Committee? | Ref: Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "43. Is a Three Lines of Defence model documented with clear accountability? | Ref: Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "44. Is an annual risk assessment conducted and reviewed by the Board? | Ref: Corp.Gov Art.17 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.8
AddHeading "4.8 AML and CFT Compliance" 2
AddBullet "45. Is a Money Laundering Reporting Officer (MLRO) appointed and registered with CBUAE? | Ref: SVF Art.14; BNPL Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "46. Is a Board-approved AML/CFT Policy and Programme in place and reviewed annually? | Ref: Federal Decree-Law No.20/2018 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "47. Is an enterprise-wide AML/CFT Risk Assessment conducted at least annually? | Ref: Cabinet Res. No.10/2019 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "48. Are CDD and EDD procedures documented and consistently applied at onboarding and ongoing? | Ref: SVF Art.14 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "49. Are customers screened against UAE, UN, OFAC, and EU sanctions lists at onboarding and ongoing? | Ref: RPSCS Art.18 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "50. Is a Transaction Monitoring System in place with defined rules and alert thresholds? | Ref: SVF Art.14; BNPL Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "51. Are Suspicious Transaction Reports filed via goAML within required timelines? | Ref: Federal Decree-Law No.20/2018 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "52. Is AML/CFT training provided to all staff annually with attendance records maintained? | Ref: SVF Art.14 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "53. Are PEP controls in place with senior management approval required for onboarding? | Ref: SVF Art.14 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.9
AddHeading "4.9 Consumer Protection" 2
AddBullet "54. Are all fees, charges, and terms clearly disclosed pre-contract for BNPL and SVF products? | Ref: BNPL Art.11; SVF Art.15 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "55. Is there a Complaints Handling Policy with defined resolution SLAs? | Ref: RPSCS Art.16; Consumer Prot.Reg | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "56. Is a complaints register maintained and reported to senior management and Board? | Ref: Consumer Prot.Reg | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "57. Are BNPL debt collection practices fair and compliant with Consumer Protection Regulation? | Ref: BNPL Art.13 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "58. Is customer data handled in compliance with UAE PDPL (Federal Decree-Law No.45/2021)? | Ref: UAE PDPL | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "59. Are vulnerable customer policies documented and applied? | Ref: Consumer Prot.Reg | Status: ___ | Evidence: ___ | Observation: ___"

# 4.10
AddHeading "4.10 IT, Cybersecurity and Data Governance" 2
AddBullet "60. Is a Board-approved Cybersecurity Policy aligned with CBUAE Cyber Risk Guidance in place? | Ref: SVF Art.17; RPSCS Art.20 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "61. Is customer and transaction data stored within the UAE (data residency requirement)? | Ref: SVF Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "62. Is data retained for a minimum of 5 years as required by CBUAE? | Ref: SVF Art.17 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "63. Are annual penetration testing and vulnerability assessments conducted by qualified parties? | Ref: RPSCS Art.20 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "64. Is there an IT and Cyber Incident Response Plan tested at least annually? | Ref: RPSCS Art.20 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "65. Are privileged access controls, MFA, and encryption standards implemented? | Ref: CBUAE Cyber Guidance | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "66. Is there a Board-approved Business Continuity Plan (BCP) and Disaster Recovery Plan (DRP)? | Ref: CBUAE BCM Standards | Status: ___ | Evidence: ___ | Observation: ___"

# 4.11
AddHeading "4.11 Outsourcing and Third-Party Management" 2
AddBullet "67. Is a Board-approved Outsourcing Policy covering risk assessment and due diligence in place? | Ref: SVF Art.19; CBUAE Outsourcing Reg | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "68. Are material outsourcing arrangements notified to CBUAE prior to execution? | Ref: SVF Art.19 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "69. Do outsourcing contracts include CBUAE audit rights, data security, and termination provisions? | Ref: CBUAE Outsourcing Reg | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "70. Is an Outsourcing Register maintained with risk ratings reviewed at least annually? | Ref: CBUAE Outsourcing Reg | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "71. Is ongoing performance monitoring of critical third-party providers documented? | Ref: CBUAE Outsourcing Reg | Status: ___ | Evidence: ___ | Observation: ___"

# 4.12
AddHeading "4.12 Regulatory Reporting and Supervisory Obligations" 2
AddBullet "72. Are all CBUAE regulatory returns (monthly, quarterly, annual) submitted accurately and on time? | Ref: SVF Art.21; RPSCS Art.25 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "73. Is there a regulatory reporting calendar owned by the Compliance team? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "74. Are material events (breaches, incidents, ownership changes) reported to CBUAE promptly? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "75. Are CBUAE inspection findings tracked and remediated within agreed timelines? | Ref: SVF Art.23 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "76. Are regulatory return submissions reviewed and approved by the CCO and CFO before filing? | Ref: SVF Art.21 | Status: ___ | Evidence: ___ | Observation: ___"

# 4.13
AddHeading "4.13 Internal Audit Function" 2
AddBullet "77. Does the Internal Audit function report independently to the Audit Committee? | Ref: Corp.Gov Art.13 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "78. Is an annual risk-based Internal Audit Plan approved by the Audit Committee? | Ref: Corp.Gov Art.13 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "79. Are internal audit findings tracked with management responses and target remediation dates? | Ref: Corp.Gov Art.13 | Status: ___ | Evidence: ___ | Observation: ___"
AddBullet "80. Is the Head of Internal Audit appointment and removal subject to Audit Committee approval? | Ref: Corp.Gov Art.13 | Status: ___ | Evidence: ___ | Observation: ___"

AddPageBreak

# SECTION 5
AddHeading "5. AUDIT FINDINGS SUMMARY AND RATING" 1
AddBody "Complete the table below after finalizing the checklist:"
AddBody ""
AddBody "Domain                             | Total | Compliant | Partially Compliant | Non-Compliant | Risk Rating"
AddBody "Board and Corporate Governance     |  10   |           |                     |               |"
AddBody "Licensing and Regulatory Standing  |   5   |           |                     |               |"
AddBody "Capital and Financial Adequacy     |   4   |           |                     |               |"
AddBody "SVF Float Safeguarding             |   5   |           |                     |               |"
AddBody "BNPL Credit Risk                   |   7   |           |                     |               |"
AddBody "Compliance Function                |   6   |           |                     |               |"
AddBody "Risk Management Framework          |   7   |           |                     |               |"
AddBody "AML and CFT Compliance             |   9   |           |                     |               |"
AddBody "Consumer Protection                |   6   |           |                     |               |"
AddBody "IT, Cybersecurity and Data         |   7   |           |                     |               |"
AddBody "Outsourcing                        |   5   |           |                     |               |"
AddBody "Regulatory Reporting               |   5   |           |                     |               |"
AddBody "Internal Audit Function            |   4   |           |                     |               |"
AddBody "TOTAL                              |  80   |           |                     |               |"

AddPageBreak

# SECTION 6
AddHeading "6. AUDIT RATING DEFINITIONS" 1
AddBullet "SATISFACTORY (90-100% Compliant): No material weaknesses; minor observations only."
AddBullet "NEEDS IMPROVEMENT (70-89% Compliant): Some control gaps; management action plan required."
AddBullet "UNSATISFACTORY (50-69% Compliant): Significant control weaknesses; immediate remediation required."
AddBullet "CRITICAL (Below 50% Compliant): Fundamental control failures; escalation to Board and possible CBUAE notification required."

AddPageBreak

# SECTION 7
AddHeading "7. DISCLAIMER" 1
AddBody "This checklist is prepared based on publicly available CBUAE regulations and standards as of April 2026. It is intended for internal audit guidance purposes only. Regulations are subject to change; auditors must verify the latest versions of all referenced regulations on the official CBUAE website at www.centralbank.ae. This document does not constitute legal or regulatory advice."

# SAVE
$savePath = "C:\Users\F2C\Downloads\CBUAE_Governance_Audit_Checklist.docx"
$doc.SaveAs([ref]$savePath, [ref]16)
$doc.Close()
$word.Quit()

Write-Host "SUCCESS: Document saved to $savePath"
