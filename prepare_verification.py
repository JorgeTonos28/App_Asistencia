import re

def combine_files():
    # Load templates
    try:
        register_html = open('Register.html').read()
        css_html = open('css.html').read()
        register_js = open('register.js.html').read()
        sig_pad_js = open('signaturepad.min.js.html').read()
    except FileNotFoundError as e:
        print(f"Error reading file: {e}")
        return

    # Mock GAS include_
    def include_replacer(match):
        fn = match.group(1)
        if fn == 'css': return css_html
        if fn == 'register.js': return register_js
        if fn == 'signaturepad.min.js': return sig_pad_js
        return ''

    content = re.sub(r'<\?!= include_\(\'(.*?)\'\); \?>', include_replacer, register_html)

    # Mock variables
    # We replace PHP-style short tags <?= ?> and <? ?> with JS or static values

    # Event data
    content = content.replace('<?= evento.nombre ?>', 'Taller de Liderazgo 2025')
    content = content.replace('<?= evento.id ?>', 'EV0001')
    content = content.replace('<?= evento.fecha ?>', '2025-10-20')
    content = content.replace('<?= logoUrl ?>', '') # Placeholder
    content = content.replace("<?= isRegistered ? 'true' : 'false' ?>", 'false') # Default not registered
    content = content.replace('<?= participantName ?>', '')

    # Simple Date replacement
    content = content.replace('<?= new Date().getFullYear() ?>', '2025')

    # PHP/GAS tags cleanup
    content = re.sub(r'<\? if \(isRegistered\) \{ \?>', '<!-- if registered -->', content)
    content = re.sub(r'<\? \} \?>', '<!-- end if -->', content)

    # Handle the "hidden" class based on isRegistered=false logic manually for the mockup
    # In the template: <? if (isRegistered) { ?> ... <? } ?>
    # Since we set isRegistered=false, we should HIDE the status card and SHOW the form.
    # The form has: class="... <? if (isRegistered) { ?>hidden<? } ?>"
    # Since isRegistered is false, the "hidden" inside the class string should NOT be rendered.

    # A simple way to mock this specific logic:
    # We replaced the surrounding tags with comments.
    # Now we need to handle the inline class condition.
    content = content.replace('<? if (isRegistered) { ?>hidden<? } ?>', '')

    # Mock google.script.run
    mock_script = """
    <script>
      // MOCK GAS ENVIRONMENT
      const google = {
        script: {
          run: {
            withSuccessHandler: function(cb) {
              this._cb = cb;
              return this;
            },
            withFailureHandler: function(cb) {
              return this;
            },
            getLogoDataUrl: function() {
              // Return a placeholder or mock
              setTimeout(() => {
                 // Use a colorful placeholder for visual verification
                 // data:image/png;base64,... is too long, let's use a dummy image logic or just null
                 if(this._cb) this._cb(null);
              }, 500);
            },
            findParticipanteByCodigo: function(cod) {
              setTimeout(() => {
                 if(this._cb) this._cb({ok:true, participante:{nombre:'Juan Perez', correo:'juan@test.com'}});
              }, 500);
            },
            registrarAsistencia: function(data) {
               setTimeout(() => {
                 if(this._cb) this._cb({ok:true});
               }, 1000);
            }
          }
        }
      };
    </script>
    """

    # Inject mock script before closing body
    content = content.replace('</body>', mock_script + '</body>')

    with open('/home/jules/verification/test_register.html', 'w') as f:
        f.write(content)
    print("Created /home/jules/verification/test_register.html")

if __name__ == "__main__":
    combine_files()
