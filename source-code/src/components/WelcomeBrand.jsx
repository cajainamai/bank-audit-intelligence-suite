import React from 'react';
import {
    FileSpreadsheet, BarChart3, Shield, TrendingUp, AlertTriangle, Lock, BookOpen, HelpCircle
} from 'lucide-react';

const FEATURES = [
    {
        icon: <Shield size={20} />,
        title: 'NPA Classification',
        desc: 'Automatic standard vs NPA account identification',
        color: 'var(--p-rose)',
        iconColor: '#9F1239'
    },
    {
        icon: <AlertTriangle size={20} />,
        title: 'Overdrawn Detection',
        desc: 'Flag accounts exceeding sanction limit by >110%',
        color: 'var(--p-peach)',
        iconColor: '#92400E'
    },
    {
        icon: <Lock size={20} />,
        title: 'Security Shortfall',
        desc: 'Identify accounts where limit exceeds primary security',
        color: 'var(--p-sky)',
        iconColor: '#075985'
    },
    {
        icon: <TrendingUp size={20} />,
        title: 'Large Advances',
        desc: 'Report exposures above the regulatory threshold',
        color: 'var(--p-lavender)',
        iconColor: '#4C1D95'
    },
    {
        icon: <BarChart3 size={20} />,
        title: 'Top 60% Advances',
        desc: 'Instant concentration analysis by individual account',
        color: 'var(--p-mint)',
        iconColor: '#064E3B'
    },
    {
        icon: <FileSpreadsheet size={20} />,
        title: 'NPA Provision Calculator',
        desc: 'Auto provision with secured/unsecured split & Excel formulas',
        color: 'var(--p-lilac)',
        iconColor: '#5B21B6'
    },
];

export default function WelcomeBrand({ onStart }) {
    return (
        <div style={{
            minHeight: '100vh',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '3rem 1.5rem',
            gap: '3rem'
        }}>
            {/* Hero */}
            <div className="animate-slide" style={{ textAlign: 'center', maxWidth: '640px' }}>
                {/* App icon */}
                <div style={{
                    width: 62, height: 62,
                    background: 'linear-gradient(135deg, var(--brand-primary), #818CF8)',
                    borderRadius: 16,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    margin: '0 auto 1.5rem',
                    boxShadow: '0 8px 24px rgba(109,100,245,0.28)'
                }}>
                    <BookOpen size={28} color="#fff" />
                </div>

                <h1 style={{
                    fontFamily: "'Inter', sans-serif",
                    fontSize: 'clamp(2rem, 5vw, 3rem)',
                    fontWeight: 800,
                    color: 'var(--text-heading)',
                    letterSpacing: '-0.03em',
                    lineHeight: 1.15,
                    marginBottom: '1rem'
                }}>
                    Bank Audit{' '}
                    <span style={{
                        background: 'linear-gradient(135deg, var(--brand-primary), var(--brand-secondary))',
                        WebkitBackgroundClip: 'text',
                        backgroundClip: 'text',
                        WebkitTextFillColor: 'transparent',
                    }}>
                        Intelligence Suite
                    </span>
                </h1>

                <p style={{
                    fontSize: '1.0625rem',
                    color: 'var(--text-muted)',
                    lineHeight: 1.7,
                    maxWidth: 500,
                    margin: '0 auto 2rem'
                }}>
                    Upload your bank data, map columns, and generate a fully formatted multi-sheet Excel audit report in minutes — with NPA classification, provision calculations, large advance analysis and more.
                </p>

                <button
                    className="btn btn-primary btn-lg"
                    onClick={onStart}
                    style={{ minWidth: 200 }}
                >
                    Start Audit Analysis →
                </button>
                <p style={{ marginTop: '0.75rem', fontSize: '0.8125rem', color: 'var(--text-faint)' }}>
                    No login required &bull; All data stays on your device
                </p>
            </div>

            {/* Feature Grid */}
            <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(210px, 1fr))',
                gap: '1rem',
                width: '100%',
                maxWidth: '800px'
            }}>
                {FEATURES.map(({ icon, title, desc, color, iconColor }, i) => (
                    <div
                        key={i}
                        className="card animate-slide"
                        style={{
                            padding: '1.125rem 1.25rem',
                            animationDelay: `${i * 0.06}s`,
                            opacity: 0
                        }}
                    >
                        <div style={{
                            width: 38, height: 38,
                            background: color,
                            borderRadius: 10,
                            display: 'flex', alignItems: 'center', justifyContent: 'center',
                            color: iconColor,
                            marginBottom: '0.75rem',
                            flexShrink: 0
                        }}>
                            {icon}
                        </div>
                        <div style={{ fontWeight: 700, fontSize: '0.875rem', color: 'var(--text-heading)', marginBottom: 4 }}>
                            {title}
                        </div>
                        <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', lineHeight: 1.5 }}>
                            {desc}
                        </div>
                    </div>
                ))}
            </div>

            {/* Pro tip */}
            <div style={{
                background: 'var(--p-sky)',
                border: '1px solid var(--p-sky-d)',
                borderRadius: 'var(--r-md)',
                padding: '0.875rem 1.25rem',
                display: 'flex',
                alignItems: 'flex-start',
                gap: 10,
                maxWidth: 640,
                width: '100%'
            }}>
                <HelpCircle size={16} color="#0369A1" style={{ marginTop: 2, flexShrink: 0 }} />
                <p style={{ fontSize: '0.8375rem', color: '#0369A1', lineHeight: 1.55 }}>
                    <strong>Optional:</strong> Upload a <strong>Product Mapping file</strong> alongside your data to auto-classify loan types (Term Loan, CC, OD, Gold, etc.) — saving manual effort.
                </p>
            </div>

            {/* Branding Footer */}
            <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: 14,
                paddingTop: '1rem',
                borderTop: '1px solid var(--border)',
                width: '100%',
                maxWidth: 640,
                justifyContent: 'center',
                flexWrap: 'wrap'
            }}>
                <div style={{
                    width: 36, height: 36,
                    background: 'linear-gradient(135deg, #6D64F5, #0EA5E9)',
                    borderRadius: 10,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    fontWeight: 800, fontSize: '0.875rem', color: '#fff',
                    fontFamily: "'Inter', sans-serif", flexShrink: 0
                }}>JS</div>
                <div>
                    <div style={{ fontWeight: 700, fontSize: '0.875rem', color: 'var(--text-heading)' }}>
                        CA Jainam Shah
                    </div>
                    <div style={{ fontSize: '0.775rem', color: 'var(--text-muted)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                        <span>Chartered Accountant</span>
                        <span style={{ color: 'var(--border-strong)' }}>·</span>
                        <a href="https://wa.me/919687070056" target="_blank" rel="noreferrer"
                            style={{ color: 'var(--brand-primary)', textDecoration: 'none', fontWeight: 500 }}>
                            +91 96870 70056
                        </a>
                        <span style={{ color: 'var(--border-strong)' }}>·</span>
                        <a href="https://www.linkedin.com/in/cajainam/" target="_blank" rel="noreferrer"
                            style={{ color: 'var(--brand-secondary)', textDecoration: 'none', fontWeight: 500 }}>
                            LinkedIn
                        </a>
                    </div>
                </div>
            </div>
        </div>
    );
}
