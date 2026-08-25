# Icon Generation

The manifest requires PNG icons at the following sizes:

- `icon-16.png` (16×16)
- `icon-32.png` (32×32)
- `icon-64.png` (64×64)
- `icon-80.png` (80×80)
- `icon-128.png` (128×128)

## Quick option — use the SVG

Convert `gorelo-logo.svg` to PNG at each size using any of:

- **Inkscape** (free): `inkscape gorelo-logo.svg -w 16 -h 16 -o icon-16.png`
- **ImageMagick**: `convert -background none gorelo-logo.svg -resize 16x16 icon-16.png`
- **Online**: https://svgtopng.com

## Or use Gorelo's own icon

If you prefer, use Gorelo's actual brand icon from their website.
The manifest just needs publicly accessible PNG URLs — you can point
`manifest.xml` at any hosted images.
