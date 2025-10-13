# Backup Cost Calculator Test Scenarios

## Test 1: Basic Validation (Current)
- Config: [1,1,1,2,1,4,100,15,15,15,15,15]
- Expected Month 1:
  - S3 Standard: 100 GB, Cost: $1.15
  - S3 Intelligent: 50 GB, Cost: $0.3125
  - Glacier IR: 50 GB, Cost: $0.10 + $0.50 penalty = $0.60
  - Total: $2.0625

## Test 2: All Tiers Used
- Config: [1,0,1,1,1,5,100,30,30,30,90,180]
- Weekly → S3 Standard, Off-account → Glacier DA
- Expected Month 1:
  - S3 Standard: (1+0+1) × 100 = 200 GB, Cost: $4.60
  - Glacier DA: 1 × 100 = 100 GB, Cost: $0.099
  - Total: $4.699

## Test 3: Zero Retention Skip
- Config: [1,1,1,2,1,4,100,0,30,90,180,365]
- S3 Standard retention = 0 → Skip hourly/daily
- Expected Month 1:
  - S3 Standard: 0 GB, Cost: $0
  - S3 Intelligent: 100 GB, Cost: $1.25
  - Glacier IR: 100 GB, Cost: $0.40
  - Total: $1.65

## Test 4: Minimum Duration Compliance
- Config: [1,1,1,3,1,4,100,30,30,90,180,365]
- Weekly → S3-IA (90 days ≥ 30 min), Off-account → Glacier IR (180 days ≥ 90 min)
- Expected Month 1: No penalties
  - S3 Standard: 200 GB, Cost: $4.60
  - S3-IA: 100 GB, Cost: $1.125
  - Glacier IR: 100 GB, Cost: $0.40
  - Total: $6.125

## Test 5: Multiple Backups Same Tier
- Config: [2,2,2,1,2,1,50,30,30,90,180,365]
- Weekly + Off-account both → S3 Standard
- Expected Month 1:
  - S3 Standard: (2+2+2+2) × 50 = 400 GB, Cost: $9.20
  - Total: $9.20

## Validation Checklist:
- [ ] Storage distribution matches tier selection
- [ ] Retention prorating works (15/30 = 0.5)
- [ ] Early deletion penalties apply correctly
- [ ] Zero retention skips storage
- [ ] Multiple backups in same tier aggregate
- [ ] Visual flow shows correct tiers and retention