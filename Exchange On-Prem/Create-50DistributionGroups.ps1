<#
    .SYNOPSIS
    Creates 50 distribution groups in Exchange 2019 On-Premises.

    .DESCRIPTION
    Script creates 50 static distribution groups with corporate-style names in Exchange 2019.
    Each group gets a primary SMTP address on the msonline.dk domain.

    .EXAMPLE
    C:\PS> .\Create-50DistributionGroups.ps1

    .NOTES
    Requires: Exchange Management Shell (Exchange 2019)

    .DISCLAIMER
    This script is provided AS-IS, with no warranty - Use at own risk.
#>

$domain = "msonline.dk"

$groups = @(
    @{ Name = "Marketing-DIST"; Alias = "marketing-dist"; DisplayName = "Marketing Distribution" }
    @{ Name = "Finance-DIST"; Alias = "finance-dist"; DisplayName = "Finance Distribution" }
    @{ Name = "HR-DIST"; Alias = "hr-dist"; DisplayName = "Human Resources Distribution" }
    @{ Name = "IT-DIST"; Alias = "it-dist"; DisplayName = "IT Distribution" }
    @{ Name = "Sales-DIST"; Alias = "sales-dist"; DisplayName = "Sales Distribution" }
    @{ Name = "Legal-DIST"; Alias = "legal-dist"; DisplayName = "Legal Distribution" }
    @{ Name = "Operations-DIST"; Alias = "operations-dist"; DisplayName = "Operations Distribution" }
    @{ Name = "Engineering-DIST"; Alias = "engineering-dist"; DisplayName = "Engineering Distribution" }
    @{ Name = "Support-DIST"; Alias = "support-dist"; DisplayName = "Support Distribution" }
    @{ Name = "Procurement-DIST"; Alias = "procurement-dist"; DisplayName = "Procurement Distribution" }
    @{ Name = "Logistics-DIST"; Alias = "logistics-dist"; DisplayName = "Logistics Distribution" }
    @{ Name = "Research-DIST"; Alias = "research-dist"; DisplayName = "Research Distribution" }
    @{ Name = "Development-DIST"; Alias = "development-dist"; DisplayName = "Development Distribution" }
    @{ Name = "QualityAssurance-DIST"; Alias = "qualityassurance-dist"; DisplayName = "Quality Assurance Distribution" }
    @{ Name = "CustomerService-DIST"; Alias = "customerservice-dist"; DisplayName = "Customer Service Distribution" }
    @{ Name = "Compliance-DIST"; Alias = "compliance-dist"; DisplayName = "Compliance Distribution" }
    @{ Name = "Accounting-DIST"; Alias = "accounting-dist"; DisplayName = "Accounting Distribution" }
    @{ Name = "ProjectManagement-DIST"; Alias = "projectmanagement-dist"; DisplayName = "Project Management Distribution" }
    @{ Name = "BusinessDev-DIST"; Alias = "businessdev-dist"; DisplayName = "Business Development Distribution" }
    @{ Name = "Communications-DIST"; Alias = "communications-dist"; DisplayName = "Communications Distribution" }
    @{ Name = "Facilities-DIST"; Alias = "facilities-dist"; DisplayName = "Facilities Distribution" }
    @{ Name = "Training-DIST"; Alias = "training-dist"; DisplayName = "Training Distribution" }
    @{ Name = "Administration-DIST"; Alias = "administration-dist"; DisplayName = "Administration Distribution" }
    @{ Name = "Executive-DIST"; Alias = "executive-dist"; DisplayName = "Executive Distribution" }
    @{ Name = "ProductMgmt-DIST"; Alias = "productmgmt-dist"; DisplayName = "Product Management Distribution" }
    @{ Name = "DataAnalytics-DIST"; Alias = "dataanalytics-dist"; DisplayName = "Data Analytics Distribution" }
    @{ Name = "CyberSecurity-DIST"; Alias = "cybersecurity-dist"; DisplayName = "Cyber Security Distribution" }
    @{ Name = "CloudOps-DIST"; Alias = "cloudops-dist"; DisplayName = "Cloud Operations Distribution" }
    @{ Name = "DevOps-DIST"; Alias = "devops-dist"; DisplayName = "DevOps Distribution" }
    @{ Name = "InternalAudit-DIST"; Alias = "internalaudit-dist"; DisplayName = "Internal Audit Distribution" }
    @{ Name = "RiskManagement-DIST"; Alias = "riskmanagement-dist"; DisplayName = "Risk Management Distribution" }
    @{ Name = "StrategicPlanning-DIST"; Alias = "strategicplanning-dist"; DisplayName = "Strategic Planning Distribution" }
    @{ Name = "SupplyChain-DIST"; Alias = "supplychain-dist"; DisplayName = "Supply Chain Distribution" }
    @{ Name = "Warehouse-DIST"; Alias = "warehouse-dist"; DisplayName = "Warehouse Distribution" }
    @{ Name = "Manufacturing-DIST"; Alias = "manufacturing-dist"; DisplayName = "Manufacturing Distribution" }
    @{ Name = "HealthSafety-DIST"; Alias = "healthsafety-dist"; DisplayName = "Health & Safety Distribution" }
    @{ Name = "EnvironmentalMgmt-DIST"; Alias = "environmentalmgmt-dist"; DisplayName = "Environmental Management Distribution" }
    @{ Name = "MediaRelations-DIST"; Alias = "mediarelations-dist"; DisplayName = "Media Relations Distribution" }
    @{ Name = "InvestorRelations-DIST"; Alias = "investorrelations-dist"; DisplayName = "Investor Relations Distribution" }
    @{ Name = "TalentAcquisition-DIST"; Alias = "talentacquisition-dist"; DisplayName = "Talent Acquisition Distribution" }
    @{ Name = "PayrollDept-DIST"; Alias = "payrolldept-dist"; DisplayName = "Payroll Department Distribution" }
    @{ Name = "BudgetPlanning-DIST"; Alias = "budgetplanning-dist"; DisplayName = "Budget Planning Distribution" }
    @{ Name = "CorporateStrategy-DIST"; Alias = "corporatestrategy-dist"; DisplayName = "Corporate Strategy Distribution" }
    @{ Name = "DigitalTransform-DIST"; Alias = "digitaltransform-dist"; DisplayName = "Digital Transformation Distribution" }
    @{ Name = "Innovation-DIST"; Alias = "innovation-dist"; DisplayName = "Innovation Distribution" }
    @{ Name = "PartnerMgmt-DIST"; Alias = "partnermgmt-dist"; DisplayName = "Partner Management Distribution" }
    @{ Name = "VendorMgmt-DIST"; Alias = "vendormgmt-dist"; DisplayName = "Vendor Management Distribution" }
    @{ Name = "ContractMgmt-DIST"; Alias = "contractmgmt-dist"; DisplayName = "Contract Management Distribution" }
    @{ Name = "PropertyMgmt-DIST"; Alias = "propertymgmt-dist"; DisplayName = "Property Management Distribution" }
    @{ Name = "FleetMgmt-DIST"; Alias = "fleetmgmt-dist"; DisplayName = "Fleet Management Distribution" }
)

Write-Output "========================================"
Write-Output "Creating 50 Distribution Groups"
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
            -Type "Distribution" `
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
