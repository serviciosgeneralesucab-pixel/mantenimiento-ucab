export const metadata = {
  title: "Coordinación de Mantenimiento",
  description: "Seguimiento de cumplimiento — rutinas, planes y correctivas",
};

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body style={{ margin: 0, background: "#F3F5F8" }}>{children}</body>
    </html>
  );
}
