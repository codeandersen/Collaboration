<#
    .SYNOPSIS
    Creates 50 mail-enabled security groups in Exchange 2019 On-Premises.

    .DESCRIPTION
    Script creates 50 mail-enabled security groups with corporate-style names in Exchange 2019.
    Each group gets a primary SMTP address on the msonline.dk domain.

    .EXAMPLE
    C:\PS> .\Create-50MailEnabledSecurityGroups.ps1

    .NOTES
    Requires: Exchange Management Shell (Exchange 2019)

    .DISCLAIMER
    This script is provided AS-IS, with no warranty - Use at own risk.
#>

$domain = "msonline.dk"

$groups = @(
    @{ Name = "Marketing-SEC"; Alias = "marketing-sec"; DisplayName = "Marketing Security" }
    @{ Name = "Finance-SEC"; Alias = "finance-sec"; DisplayName = "Finance Security" }
    @{ Name = "HR-SEC"; Alias = "hr-sec"; DisplayName = "Human Resources Security" }
    @{ Name = "IT-SEC"; Alias = "it-sec"; DisplayName = "IT Security" }
    @{ Name = "Sales-SEC"; Alias = "sales-sec"; DisplayName = "Sales Security" }
    @{ Name = "Legal-SEC"; Alias = "legal-sec"; DisplayName = "Legal Security" }
    @{ Name = "Operations-SEC"; Alias = "operations-sec"; DisplayName = "Operations Security" }
    @{ Name = "Engineering-SEC"; Alias = "engineering-sec"; DisplayName = "Engineering Security" }
    @{ Name = "Support-SEC"; Alias = "support-sec"; DisplayName = "Support Security" }
    @{ Name = "Procurement-SEC"; Alias = "procurement-sec"; DisplayName = "Procurement Security" }
    @{ Name = "Logistics-SEC"; Alias = "logistics-sec"; DisplayName = "Logistics Security" }
    @{ Name = "Research-SEC"; Alias = "research-sec"; DisplayName = "Research Security" }
    @{ Name = "Development-SEC"; Alias = "development-sec"; DisplayName = "Development Security" }
    @{ Name = "QualityAssurance-SEC"; Alias = "qualityassurance-sec"; DisplayName = "Quality Assurance Security" }
    @{ Name = "CustomerService-SEC"; Alias = "customerservice-sec"; DisplayName = "Customer Service Security" }
    @{ Name = "Compliance-SEC"; Alias = "compliance-sec"; DisplayName = "Compliance Security" }
    @{ Name = "Accounting-SEC"; Alias = "accounting-sec"; DisplayName = "Accounting Security" }
    @{ Name = "ProjectManagement-SEC"; Alias = "projectmanagement-sec"; DisplayName = "Project Management Security" }
    @{ Name = "BusinessDev-SEC"; Alias = "businessdev-sec"; DisplayName = "Business Development Security" }
    @{ Name = "Communications-SEC"; Alias = "communications-sec"; DisplayName = "Communications Security" }
    @{ Name = "Facilities-SEC"; Alias = "facilities-sec"; DisplayName = "Facilities Security" }
    @{ Name = "Training-SEC"; Alias = "training-sec"; DisplayName = "Training Security" }
    @{ Name = "Administration-SEC"; Alias = "administration-sec"; DisplayName = "Administration Security" }
    @{ Name = "Executive-SEC"; Alias = "executive-sec"; DisplayName = "Executive Security" }
    @{ Name = "ProductMgmt-SEC"; Alias = "productmgmt-sec"; DisplayName = "Product Management Security" }
    @{ Name = "DataAnalytics-SEC"; Alias = "dataanalytics-sec"; DisplayName = "Data Analytics Security" }
    @{ Name = "CyberSecurity-SEC"; Alias = "cybersecurity-sec"; DisplayName = "Cyber Security" }
    @{ Name = "CloudOps-SEC"; Alias = "cloudops-sec"; DisplayName = "Cloud Operations Security" }
    @{ Name = "DevOps-SEC"; Alias = "devops-sec"; DisplayName = "DevOps Security" }
    @{ Name = "InternalAudit-SEC"; Alias = "internalaudit-sec"; DisplayName = "Internal Audit Security" }
    @{ Name = "RiskManagement-SEC"; Alias = "riskmanagement-sec"; DisplayName = "Risk Management Security" }
    @{ Name = "StrategicPlanning-SEC"; Alias = "strategicplanning-sec"; DisplayName = "Strategic Planning Security" }
    @{ Name = "SupplyChain-SEC"; Alias = "supplychain-sec"; DisplayName = "Supply Chain Security" }
    @{ Name = "Warehouse-SEC"; Alias = "warehouse-sec"; DisplayName = "Warehouse Security" }
    @{ Name = "Manufacturing-SEC"; Alias = "manufacturing-sec"; DisplayName = "Manufacturing Security" }
    @{ Name = "HealthSafety-SEC"; Alias = "healthsafety-sec"; DisplayName = "Health & Safety Security" }
    @{ Name = "EnvironmentalMgmt-SEC"; Alias = "environmentalmgmt-sec"; DisplayName = "Environmental Management Security" }
    @{ Name = "MediaRelations-SEC"; Alias = "mediarelations-sec"; DisplayName = "Media Relations Security" }
    @{ Name = "InvestorRelations-SEC"; Alias = "investorrelations-sec"; DisplayName = "Investor Relations Security" }
    @{ Name = "TalentAcquisition-SEC"; Alias = "talentacquisition-sec"; DisplayName = "Talent Acquisition Security" }
    @{ Name = "PayrollDept-SEC"; Alias = "payrolldept-sec"; DisplayName = "Payroll Department Security" }
    @{ Name = "BudgetPlanning-SEC"; Alias = "budgetplanning-sec"; DisplayName = "Budget Planning Security" }
    @{ Name = "CorporateStrategy-SEC"; Alias = "corporatestrategy-sec"; DisplayName = "Corporate Strategy Security" }
    @{ Name = "DigitalTransform-SEC"; Alias = "digitaltransform-sec"; DisplayName = "Digital Transformation Security" }
    @{ Name = "Innovation-SEC"; Alias = "innovation-sec"; DisplayName = "Innovation Security" }
    @{ Name = "PartnerMgmt-SEC"; Alias = "partnermgmt-sec"; DisplayName = "Partner Management Security" }
    @{ Name = "VendorMgmt-SEC"; Alias = "vendormgmt-sec"; DisplayName = "Vendor Management Security" }
    @{ Name = "ContractMgmt-SEC"; Alias = "contractmgmt-sec"; DisplayName = "Contract Management Security" }
    @{ Name = "PropertyMgmt-SEC"; Alias = "propertymgmt-sec"; DisplayName = "Property Management Security" }
    @{ Name = "FleetMgmt-SEC"; Alias = "fleetmgmt-sec"; DisplayName = "Fleet Management Security" }
)

Write-Output "========================================"
Write-Output "Creating 50 Mail-Enabled Security Groups"
Write-Output "Domain: $domain"
Write-Output "========================================"
Write-Output ""

$successCount = 0
$failCount = 0

foreach ($group in $groups) {
    $primarySmtp = "$($group.Alias)@$domain"
    
    try {
        New-DistributionGroup `
            -Name $group.Name `
            -Alias $group.Alias `
            -DisplayName $group.DisplayName `
            -PrimarySmtpAddress $primarySmtp `
            -Type "Security" `
            -ErrorAction Stop

        Write-Output "[+] Created: $($group.Name) ($primarySmtp)"
        $successCount++
    }
    catch {
        Write-Warning "[!] Failed to create $($group.Name): $($_.Exception.Message)"
        $failCount++
    }
}

Write-Output ""
Write-Output "========================================"
Write-Output "Summary"
Write-Output "========================================"
Write-Output "Successfully created: $successCount"
Write-Output "Failed: $failCount"
Write-Output "========================================"
