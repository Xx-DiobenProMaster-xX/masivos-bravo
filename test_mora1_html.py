import unittest
from mora1_html import html_mora1

class RenderTest(unittest.TestCase):
    def test_markdown(self):
        result=html_mora1("Hola **cliente**.\n\n**Línea de atención:** WhatsApp")
        self.assertIn("<strong>cliente</strong>",result)
        self.assertIn("<strong>Línea de atención:</strong>",result)
        self.assertIn("<table",result)
        self.assertNotIn("**",result)
    def test_html_existing(self):
        content="<html><body><p>Original</p></body></html>"
        self.assertEqual(html_mora1(content),content)
    def test_html_escape(self):
        self.assertIn("&lt;script&gt;",html_mora1("Hola <script>"))
if __name__=="__main__":unittest.main()
