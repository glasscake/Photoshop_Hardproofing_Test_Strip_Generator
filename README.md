# Photoshop_Hardproofing_Test_Strip_Generator
a photoshop script that generates non destructive test strips for hardproofing

AI Warning, this is a completely vibe coded app, very minimal human interaction outside of testing the functionality was put into making this

To use simply crop a strip you would like to generate sample strips from. This crop should be its own document with no background layer and the crop should be unlocked.
Select the layer youd like to generate the test strips from then go to File -> Scripts -> Browse and select this script.
Select what you would like to sweep, its starting point, and how many steps in each direction you would like to do.
It should then generate them in a new document that can then be printed.

You can then copy back the mask you picked to the cropped document and stack additional filters ontop of that. My normal workflow is finding the best exposure then copying it back and finding the best contrast. This feature should work for as many filters you have. Unexpected behavior may happen if you stack multiple of the same filter.

<img src="demo.gif">
