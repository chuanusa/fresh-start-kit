const NotFound = () => {
  return (
    <div className="min-h-screen flex items-center justify-center bg-background">
      <div className="text-center">
        <h1 className="text-6xl font-bold text-primary mb-4">404</h1>
        <p className="text-muted-foreground mb-6">找不到此頁面</p>
        <a href="/" className="text-primary hover:underline">回到首頁</a>
      </div>
    </div>
  );
};

export default NotFound;
