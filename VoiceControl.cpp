/*
  Copyright (C) 2015 David Grace

  This program is free software; you can redistribute it and/or
  modify it under the terms of the GNU General Public License
  as published by the Free Software Foundation; either version 2
  of the License, or (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program; if not, write to the Free Software
  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

#include <Windows.h>
//#include <cstdio>
#include <vector>
#include <cassert>
#include <sphelper.h>

#include "../../SDK/API/RainmeterAPI.h"

// Overview: MIDI message reception and parsing.

// Sample skin:
/*
	[mVC_Parent]
	Measure=Plugin
	Plugin=VoiceControl.dll
	GrammarFile=@VC.grxml

	[mVC_Activate]
	Measure=Plugin
	Plugin=VoiceControl.dll
	Parent=mVC_Parent
	Keyword="Rainmeter"
*/


#define EXIT_ON_ERROR(hres)		if(FAILED(hres)) { goto Exit; }
#define GID_DICTATION			0									// Dictation grammar has grammar ID 0
#define WM_RECOEVENT			WM_APP								// Window message used for recognition events


/**
 * Container for the recognizer system setup.
 */
struct Device
{
	CComPtr<ISpRecognizer>		cpRecognizer;
	CComPtr<ISpRecoContext>		cpContext;
	CComPtr<ISpAudio>			cpAudio;
	CComPtr<ISpRecoGrammar>		cpGrammar;
	bool						m_sndActive;
	int							m_nReco;
	WCHAR						m_reco[8][256];

	Device ()
		: m_sndActive	(false)
		, m_nReco		(0)
	{
		memset(m_reco, 0, sizeof(m_reco));
	}
};


/**
 * A Rainmeter measure - can be either a parent, which initializes the hardware recognizer
 * Device, or a child, which just references the parent with a specific trigger keyword.
 */
struct Measure
{
	Measure*		m_parent;			///< parent measure, if any
	void*			m_skin;				///< skin pointer
	LPCWSTR			m_rmName;			///< measure name
	Device*			m_device;			///< if a parent measure, this points to the recognizer system device
	LPCWSTR			m_keyword;			///< trigger keyword

			Measure()
				: m_parent		(NULL)
				, m_skin		(NULL)
				, m_rmName		(NULL)
				, m_device		(NULL)
				, m_keyword		(NULL)
			{}
	HRESULT	DeviceInit		(void* rm, LPCWSTR fnGrammar);
	void	DeviceRelease	();
};


std::vector<Measure*> s_parents;


/**
 * Create and initialize a measure instance.  Creates WASAPI loopback
 * device if not a child measure.
 *
 * @param[out]	data			Pointer address in which to return measure instance.
 * @param[in]	rm				Rainmeter context.
 */
PLUGIN_EXPORT void Initialize (void** data, void* rm)
{
	Measure* m	= new Measure;
	m->m_skin	= RmGetSkin(rm);
	m->m_rmName	= RmGetMeasureName(rm);
	*data		= m;

	// parse parent specifier, if appropriate
	LPCWSTR parentName = RmReadString(rm, L"Parent", L"");
	if(*parentName) {
		// match parent using measure name and skin handle
		for(
				std::vector<Measure*>::const_iterator iter = s_parents.begin();
				iter != s_parents.end();
				++iter
		) {
			if(
					!_wcsicmp((*iter)->m_rmName, parentName)
					&& (*iter)->m_skin == m->m_skin
					&& !(*iter)->m_parent
			) {
				m->m_parent = (*iter);
				return;
			}
		}
		RmLogF(rm, LOG_ERROR, L"Couldn't find Parent measure '%s'.\n", parentName);
	}

	// this is a parent measure - add it to the global list
	s_parents.push_back(m);

	// parse grammar filename (optional)
	LPCWSTR fnGrammar = RmReadString(rm, L"GrammarFile", L"");

	// initialize the recognizer system
	m->DeviceInit(rm, fnGrammar);
}


/**
 * Destroy the measure instance.
 *
 * @param[in]	data			Measure instance pointer.
 */
PLUGIN_EXPORT void Finalize (void* data)
{
	Measure* m = (Measure*)data;

	m->DeviceRelease();

	if(!m->m_parent) {
		std::vector<Measure*>::iterator iter = std::find(s_parents.begin(), s_parents.end(), m);
		s_parents.erase(iter);
	}

	delete m;
}


/**
 * (Re-)parse parameters from .ini file.
 *
 * @param[in]	data			Measure instance pointer.
 * @param[in]	rm				Rainmeter context.
 * @param[out]	maxValue		?
 */
PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	Measure*	m		= (Measure*)data;

	// parse keyword specifier
	m->m_keyword		= RmReadString(rm, L"Keyword", L"");
}


/**
 * Update the measure.
 *
 * @param[in]	data			Measure instance pointer.
 * @return						1.0 if keyword event triggered since last update, else 0.
 */
PLUGIN_EXPORT double Update (void* data)
{
	Measure*	m		= (Measure*)data;
	Measure*	parent	= (m->m_parent)? m->m_parent : m;
	Device*		dev		= m->m_device;

	// poll messages if a parent measure
	if(dev) {
		CSpEvent event;
		while(event.GetFrom(dev->cpContext) == S_OK) {
			switch(event.eEventId) {

			case SPEI_SOUND_START:
				dev->m_nReco		= 0;
				dev->m_sndActive	= TRUE;
				break;

			case SPEI_SOUND_END:
				if(dev->m_sndActive && !dev->m_nReco) {
					// the sound has started and ended, but the engine has not succeeded in recognizing anything
				}
				dev->m_sndActive	= false;
				break;

			case SPEI_RECOGNITION:
				// there may be multiple recognition results, so get all of them
				{
					dev->m_nReco	= 0;
					CSpDynamicString dstrText;
					if(
							(dev->m_nReco < 8)
							&& !FAILED(event.RecoResult()->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &dstrText, NULL))
					) {
						_snwprintf_s(
								dev->m_reco[dev->m_nReco++],
								_TRUNCATE,
								L"%s",
								(LPTSTR)CW2T(dstrText)
						);
					}
				}
				break;

			}
		}
		return dev->m_sndActive;
	}

	return 0.0;
}


/**
 * Get a string value from the measure.
 *
 * @param[in]	data			Measure instance pointer.
 * @return						String value - must be copied out by the caller.
 */
PLUGIN_EXPORT LPCWSTR GetString (void* data)
{
	static WCHAR	buf[1024];
	Measure*		m		= (Measure*)data;
	Device*			dev		= m->m_device;

	if(dev) {
		WCHAR*		d		= buf;
		*d					= '\0';
		for(int iReco=0; iReco<dev->m_nReco; ++iReco) {
			WCHAR*	s		= dev->m_reco[iReco];
			d				+= _snwprintf_s(
									d,
									(sizeof(buf)+(UINT32)(buf-d))/sizeof(WCHAR),
									_TRUNCATE,
									L"%s%s",
									(iReco > 0)? L" " : L"",
									s
							);
		}
		return buf;
	}

	return L"";
}


#if 0
/**
 * Send a command to the measure.
 *
 * @param[in]	data			Measure instance pointer.
 * @param[in]	args			Command arguments.
 */
PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
	Measure*	m		= (Measure*)data;
}
#endif


/**
 * Try to initialize the speech recognizer system.
 * See: https://msdn.microsoft.com/en-us/library/jj127591
 *
 * @return		Result value, S_OK on success.
 */
HRESULT	Measure::DeviceInit (void* rm, LPCWSTR fnGrammar)
{
	HRESULT	hr	= S_OK;
	RmLogF(rm, LOG_DEBUG, L"Initializing speech recognizer device.\n");

	assert(!m_device);		// device already created?
	m_device	= new Device();

	hr = m_device->cpRecognizer.CoCreateInstance(CLSID_SpInprocRecognizer);
	EXIT_ON_ERROR(hr)

	hr = m_device->cpRecognizer->CreateRecoContext(&m_device->cpContext);
	EXIT_ON_ERROR(hr)

	// Set recognition notification for dictation
	hr = m_device->cpContext->SetNotifyWindowMessage(RmGetSkinWindow(rm), WM_RECOEVENT, 0, 0);
	EXIT_ON_ERROR(hr)

	// This specifies which of the recognition events are going to trigger notifications.
	// Here, all we are interested in is the beginning and ends of sounds, as well as
	// when the engine has recognized something
	const ULONGLONG ullInterest = SPFEI(SPEI_RECOGNITION) | SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_SOUND_END);
	hr = m_device->cpContext->SetInterest(ullInterest, ullInterest);
	EXIT_ON_ERROR(hr)

	// create default audio object
	hr = SpCreateDefaultObjectFromCategoryId(SPCAT_AUDIOIN, &m_device->cpAudio);
	EXIT_ON_ERROR(hr)

	// set the input for the engine
	hr = m_device->cpRecognizer->SetInput(m_device->cpAudio, TRUE);
	EXIT_ON_ERROR(hr)

	// Initialize the grammar
	if(fnGrammar && *fnGrammar) {
		// load grammar from a file
		hr = m_device->cpContext->CreateGrammar(0, &m_device->cpGrammar );
		EXIT_ON_ERROR(hr)

		hr = m_device->cpGrammar->LoadCmdFromFile(fnGrammar, SPLO_STATIC);
		if(FAILED(hr)) {
			RmLogF(rm, LOG_ERROR, L"Error loading grammar file '%s' (code: 0x%08x).  See https://msdn.microsoft.com/en-us/library/jj127916.aspx for file format info.\n", fnGrammar, hr);
		}
		EXIT_ON_ERROR(hr)

		hr = m_device->cpGrammar->SetRuleState(NULL, NULL, SPRS_ACTIVE);
		EXIT_ON_ERROR(hr)
	} else {
		// setup for dictation mode
		hr = m_device->cpContext->CreateGrammar( GID_DICTATION, &m_device->cpGrammar );
		EXIT_ON_ERROR(hr)

		hr = m_device->cpGrammar->LoadDictation(NULL, SPLO_STATIC);
		EXIT_ON_ERROR(hr)

		hr = m_device->cpGrammar->SetDictationState( SPRS_ACTIVE );
		EXIT_ON_ERROR(hr)
	}

	hr = m_device->cpRecognizer->SetRecoState( SPRST_ACTIVE_ALWAYS );
	EXIT_ON_ERROR(hr)

	return S_OK;

Exit:
	DeviceRelease();
	return hr;
}


/**
 * Release speech recognizer system resources.
 */
void Measure::DeviceRelease ()
{
	if(m_device) {
		RmLog(LOG_DEBUG, L"Releasing speech recognizer device.\n");

		m_device->cpGrammar.Release();
		m_device->cpAudio.Release();
		m_device->cpContext.Release();
		m_device->cpRecognizer.Release();
		delete m_device;
		m_device = NULL;
	}
}
