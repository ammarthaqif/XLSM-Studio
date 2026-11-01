export default function handler(req: any, res: any) {
  res.setHeader('Content-Type', 'application/json');
  res.statusCode = 200;
  res.end(
    JSON.stringify({
      status: 'ok',
      service: 'XLSM Cloud Studio',
      timestamp: new Date().toISOString(),
      version: '1.0.0',
      author: 'Ammar Thaqif'
    })
  );
}
