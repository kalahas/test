#!/usr/bin/env node

const path = require( 'path' );
const PptxGenJS = require( 'pptxgenjs' );

const DEFAULT_OUTPUT = 'copilot-deck.pptx';
const DEFAULT_MODEL = process.env.COPILOT_MODEL || 'gpt-5';

function parseArgs( args ) {
	const promptParts = [];
	let outputPath = DEFAULT_OUTPUT;
	let model = DEFAULT_MODEL;

	for ( let i = 0; i < args.length; i++ ) {
		const arg = args[ i ];

		if ( arg === '--out' || arg === '-o' ) {
			outputPath = args[ i + 1 ] || outputPath;
			i++;
			continue;
		}

		if ( arg === '--model' ) {
			model = args[ i + 1 ] || model;
			i++;
			continue;
		}

		promptParts.push( arg );
	}

	return {
		model,
		outputPath,
		prompt: promptParts.join( ' ' ).trim(),
	};
}

function parseSlides( content ) {
	const normalizedContent = typeof content === 'string' ? content : JSON.stringify( content );
	const fences = normalizedContent.match( /```(?:json)?\s*([\s\S]*?)```/i );
	const jsonText = fences ? fences[ 1 ] : normalizedContent;
	const parsed = JSON.parse( jsonText );

	if ( ! Array.isArray( parsed ) || ! parsed.length ) {
		throw new Error( 'Copilot response did not contain a non-empty JSON array of slides.' );
	}

	return parsed;
}

async function getSlidesFromCopilot( prompt, model ) {
	const { CopilotClient, approveAll } = await import( '@github/copilot-sdk' );
	const client = new CopilotClient( {
		gitHubToken: process.env.GITHUB_TOKEN,
		cwd: process.cwd(),
	} );

	await client.start();
	const session = await client.createSession( {
		model,
		onPermissionRequest: approveAll,
	} );

	try {
		const event = await session.sendAndWait( {
			prompt: `${ prompt }

Return only JSON (no markdown) with this exact shape:
[
  {
    "title": "Slide title",
    "bullets": [ "bullet 1", "bullet 2" ]
  }
]

Use 5 slides maximum.`,
		} );

		if ( ! event || ! event.data || ! event.data.content ) {
			throw new Error( 'No assistant response received from Copilot.' );
		}

		return parseSlides( event.data.content );
	} finally {
		await session.disconnect();
		await client.stop();
	}
}

function createDeck( slides, outputFile ) {
	const pptx = new PptxGenJS();
	pptx.layout = 'LAYOUT_WIDE';
	pptx.author = 'GitHub Copilot SDK sample';
	pptx.subject = 'Generated from prompt';
	pptx.title = 'Copilot-generated deck';

	slides.forEach( ( slideData ) => {
		const slide = pptx.addSlide();
		const title = slideData && slideData.title ? String( slideData.title ) : 'Untitled';
		const bullets = Array.isArray( slideData && slideData.bullets ) ? slideData.bullets : [];

		slide.addText( title, {
			x: 0.5,
			y: 0.3,
			w: 12.3,
			h: 0.8,
			fontSize: 30,
			bold: true,
			color: '1F2937',
		} );

		if ( bullets.length ) {
			slide.addText(
				bullets.map( ( bullet ) => ( { text: String( bullet ), options: { bullet: { indent: 18 } } } ) ),
				{
					x: 0.8,
					y: 1.3,
					w: 11.5,
					h: 5,
					fontSize: 20,
					color: '111827',
					breakLine: true,
				}
			);
		}
	} );

	return pptx.writeFile( { fileName: outputFile } );
}

async function main() {
	const args = process.argv.slice( 2 );
	const parsed = parseArgs( args );
	const outputPathArg = parsed.outputPath;
	const model = parsed.model;
	const prompt = parsed.prompt;

	if ( ! prompt ) {
		console.error( 'Usage: npm run copilot:ppt -- "<prompt>" [--out file.pptx] [--model model-name]' );
		process.exitCode = 1;
		return;
	}

	const outputFile = path.resolve( process.cwd(), outputPathArg );
	const slides = await getSlidesFromCopilot( prompt, model );
	await createDeck( slides, outputFile );
	console.log( `Created ${ outputFile }` );
}

main().catch( ( error ) => {
	console.error( error.message || error );
	process.exitCode = 1;
} );
