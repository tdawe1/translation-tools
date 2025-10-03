import Navigation from '@/components/layout/Navigation';

export default function DashboardLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <div className="min-h-screen bg-gray-50">
      <Navigation />
      <div className="lg:pl-64">
        {/* Mobile spacer */}
        <div className="lg:hidden h-16" />
        <main>{children}</main>
      </div>
    </div>
  );
}