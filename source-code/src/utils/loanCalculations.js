import { differenceInMonths, isValid, getYear, getMonth } from 'date-fns';

/**
 * Calculates the monthly EMI
 * Formula: P * R * (1+R)^N / ((1+R)^N - 1)
 */
export const calculateEMI = (principal, annualRate, tenureMonths) => {
    if (!principal || !annualRate || !tenureMonths) return 0;

    const r = (annualRate / 100) / 12; // monthly rate
    const n = tenureMonths;

    if (r === 0) return principal / n;

    const emi = (principal * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1);
    return Math.round(emi);
};

/**
 * Calculates Theoretical Outstanding
 * Outstanding after n instalments are paid
 */
export const calculateTheoOutstanding = (principal, annualRate, tenureMonths, instalmentsServiced) => {
    if (!principal || !tenureMonths) return principal || 0;

    const r = (annualRate / 100) / 12;
    const N = tenureMonths;
    const n = Math.min(instalmentsServiced, N); // Can't service more than N

    if (n <= 0) return principal;
    if (n >= N) return 0;
    if (r === 0) return Math.max(0, principal - (principal / N) * n);

    const outstanding = principal * (Math.pow(1 + r, N) - Math.pow(1 + r, n)) / (Math.pow(1 + r, N) - 1);
    return Math.max(0, Math.round(outstanding));
};

export const getInstalmentsDue = (startDate, auditDate) => {
    if (!isValid(startDate) || !isValid(auditDate)) return 0;
    if (startDate > auditDate) return 0;

    // Inclusive counting: (Years * 12) + Months + 1
    const years = getYear(auditDate) - getYear(startDate);
    const months = getMonth(auditDate) - getMonth(startDate);

    return Math.max(0, (years * 12) + months + 1);
};

/**
 * Calculates interest and principal bifurcation using iterative amortization
 */
export const getAmortizationBreakdown = (principal, annualRate, tenureMonths, instalmentsDue) => {
    if (!principal || !tenureMonths) return { principalPaid: 0, interestCharged: 0, theoBalance: principal || 0 };

    const r = (annualRate / 100) / 12;
    const N = tenureMonths;
    const n = Math.min(Math.max(0, instalmentsDue), N);
    const emi = calculateEMI(principal, annualRate, N);

    if (r === 0) {
        const principalPerMonth = principal / N;
        const paid = principalPerMonth * n;
        return {
            principalPaid: Math.round(paid),
            interestCharged: 0,
            theoBalance: Math.round(principal - paid)
        };
    }

    let currentBalance = principal;
    let totalInterest = 0;
    let totalPrincipal = 0;

    for (let i = 0; i < n; i++) {
        const interest = currentBalance * r;
        const principalPart = emi - interest;

        totalInterest += interest;
        totalPrincipal += principalPart;
        currentBalance -= principalPart;
    }

    return {
        principalPaid: Math.round(totalPrincipal),
        interestCharged: Math.round(totalInterest),
        theoBalance: Math.max(0, Math.round(currentBalance))
    };
};

export const calculateTenure = (sanctionDate, expiryDate) => {
    if (!isValid(sanctionDate) || !isValid(expiryDate)) return 0;
    if (sanctionDate >= expiryDate) return 0;
    return differenceInMonths(expiryDate, sanctionDate);
};

export const calculateIrregularity = (cbsOutstanding, theoOutstanding, emi) => {
    // Strictly use Method A (Outstanding Difference)
    if (cbsOutstanding > theoOutstanding && emi > 0) {
        const diff = cbsOutstanding - theoOutstanding;
        return Math.max(0, Math.floor(diff / emi));
    }

    return 0;
};

export const getSMATag = (irregularityMonths) => {
    if (irregularityMonths <= 1) return { tag: 'Standard', level: 0 };
    if (irregularityMonths <= 2) return { tag: 'SMA-0', level: 1 };
    if (irregularityMonths <= 3) return { tag: 'SMA-1', level: 2 };
    if (irregularityMonths <= 4) return { tag: 'SMA-2', level: 3 };
    return { tag: 'NPA', level: 4 };
};
