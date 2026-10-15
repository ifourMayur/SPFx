import {
  FEATURE_CODES,
  FEATURE_NAMES,
  IFeatureSubscription,
  isFeatureCode,
  isMarketplaceReturn,
  isNavigableUrl,
  MARKETPLACE_TOKEN_PARAMETER,
  needsMarketplaceCheckout,
  offerUrlFor,
  readMarketplaceToken,
  toActivatedFeatureSubscription,
  toFeatureSubscription,
  toFeatureSubscriptions,
  toPlanLabel
} from './FeatureSubscription';

describe('toFeatureSubscription', () => {
  it('maps an active subscription', () => {
    expect(
      toFeatureSubscription(
        { feature: 'ED_CONTROL', isActive: true, subscriptionStatus: 'Active', planId: 'EDCONTROL_BASIC' },
        'ED_CONTROL'
      )
    ).toEqual({
      feature: 'ED_CONTROL',
      state: 'active',
      isActive: true,
      isEnabled: true,
      subscriptionStatus: 'Active',
      planId: 'EDCONTROL_BASIC',
      planName: undefined
    });
  });

  // Paid for and switched off: the subscription stands, the feature is not usable.
  it('tells a paid-for feature that is switched off apart from an unusable one', () => {
    const subscription: IFeatureSubscription | undefined = toFeatureSubscription(
      { feature: 'ED_CONTROL', isActive: false, isEnabled: false, subscriptionStatus: 'Active' },
      'ED_CONTROL'
    );

    expect(subscription?.state).toBe('active');
    expect(subscription?.isActive).toBe(false);
    expect(subscription?.isEnabled).toBe(false);
  });

  // Switched on is a choice, not a purchase.
  it('keeps a switched-on feature that was never bought unusable', () => {
    const subscription: IFeatureSubscription | undefined = toFeatureSubscription(
      { feature: 'ED_CONTROL', isActive: false, isEnabled: true, subscriptionStatus: 'Inactive' },
      'ED_CONTROL'
    );

    expect(subscription?.isEnabled).toBe(true);
    expect(subscription?.isActive).toBe(false);
    expect(subscription?.state).toBe('inactive');
  });

  it('reads a usable feature as switched on even when an older API leaves the switch out', () => {
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: true }, 'ED_CONTROL')?.isEnabled).toBe(true);
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: false }, 'ED_CONTROL')?.isEnabled).toBe(false);
  });

  it('maps an inactive subscription', () => {
    const subscription: IFeatureSubscription | undefined = toFeatureSubscription(
      { feature: 'ED_CONTROL', isActive: false, subscriptionStatus: 'Inactive', requiresSubscription: true },
      'ED_CONTROL'
    );

    expect(subscription?.state).toBe('inactive');
    expect(subscription?.isActive).toBe(false);
  });

  // `isActive` is the authority. The status text only says which kind of "not active".
  it('is not active on the status text alone', () => {
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: false, subscriptionStatus: 'Active' }, 'ED_CONTROL')?.isActive).toBe(false);
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: null, subscriptionStatus: 'Active' }, 'ED_CONTROL')?.isActive).toBe(false);
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', subscriptionStatus: 'Subscribed' }, 'ED_CONTROL')?.isActive).toBe(false);
  });

  it('takes only an explicit true as active', () => {
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: 'true' as unknown as boolean }, 'ED_CONTROL')?.isActive).toBe(false);
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: 1 as unknown as boolean }, 'ED_CONTROL')?.isActive).toBe(false);
  });

  it.each([
    ['Subscribed', 'active'],
    ['Pending', 'pending'],
    ['PendingFulfillmentStart', 'pending'],
    ['processing', 'pending'],
    ['Suspended', 'suspended'],
    ['Unsubscribed', 'inactive'],
    ['', 'inactive']
  ])('reads status %p as %p', (status: string, state: string) => {
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: false, subscriptionStatus: status }, 'ED_CONTROL')?.state).toBe(state);
  });

  // A mis-routed or mis-cached answer about one feature must never unlock another.
  it('rejects an answer about a different feature', () => {
    expect(toFeatureSubscription({ feature: 'ED_CONTROL', isActive: true }, 'MINUTES')).toBeUndefined();
  });

  it('accepts the feature name in any case', () => {
    expect(toFeatureSubscription({ feature: 'ed_control', isActive: true }, 'ED_CONTROL')?.isActive).toBe(true);
  });

  // The API always names the feature. A body that does not - a proxy's `{}`, say - read as
  // "not bought" would invite a second purchase; it is "could not verify" instead.
  it('rejects an answer that does not name the feature', () => {
    expect(toFeatureSubscription({}, 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription({ isActive: true }, 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription({ feature: null, isActive: false }, 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription({ feature: '  ', isActive: false }, 'ED_CONTROL')).toBeUndefined();
  });

  it('rejects a payload that is not an answer at all', () => {
    expect(toFeatureSubscription(undefined, 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription(null, 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription('Active', 'ED_CONTROL')).toBeUndefined();
    expect(toFeatureSubscription([{ isActive: true }], 'ED_CONTROL')).toBeUndefined();
  });

  it('drops blank plan fields', () => {
    const subscription = toFeatureSubscription({ feature: 'KYP', isActive: true, planId: '  ', planName: null }, 'KYP');

    expect(subscription?.planId).toBeUndefined();
    expect(subscription?.planName).toBeUndefined();
  });
});

describe('toFeatureSubscriptions', () => {
  it('keeps each feature its own answer', () => {
    const subscriptions: IFeatureSubscription[] = toFeatureSubscriptions([
      { feature: 'ED_CONTROL', isActive: true, subscriptionStatus: 'Active' },
      { feature: 'MINUTES', isActive: false, subscriptionStatus: 'Inactive' }
    ]);

    expect(subscriptions.map((item: IFeatureSubscription) => [item.feature, item.isActive])).toEqual([
      ['ED_CONTROL', true],
      ['MINUTES', false]
    ]);
  });

  it('drops features this build does not know, and entries that do not parse', () => {
    expect(toFeatureSubscriptions([{ feature: 'TELEPORT', isActive: true }, null, 'x', { isActive: true }])).toEqual([]);
  });

  it('reads anything but an array as none', () => {
    expect(toFeatureSubscriptions({ feature: 'KYP', isActive: true })).toEqual([]);
    expect(toFeatureSubscriptions(undefined)).toEqual([]);
  });
});

describe('feature codes', () => {
  it('is exactly the four sold features', () => {
    expect([...FEATURE_CODES]).toEqual(['ED_CONTROL', 'KYP', 'MINUTES', 'TENDER']);
    expect(FEATURE_NAMES).toEqual({
      ED_CONTROL: 'ED Control',
      KYP: 'KYP',
      MINUTES: 'Minutes of Meetings',
      TENDER: 'Tender'
    });
    expect(isFeatureCode('BIM')).toBe(false);
    expect(isFeatureCode('REVIEW')).toBe(false);
  });

  it('names every feature', () => {
    FEATURE_CODES.forEach((code) => expect(FEATURE_NAMES[code]).toBeTruthy());
  });

  it('recognises codes exactly', () => {
    expect(isFeatureCode('ED_CONTROL')).toBe(true);
    expect(isFeatureCode('ed_control')).toBe(false);
    expect(isFeatureCode(undefined)).toBe(false);
  });
});

describe('toPlanLabel', () => {
  it('prefers the plan name over its id', () => {
    const base: IFeatureSubscription = {
      feature: 'ED_CONTROL',
      state: 'active',
      isActive: true,
      isEnabled: true,
      subscriptionStatus: 'Active',
      planId: 'EDCONTROL_PRO'
    };

    expect(toPlanLabel({ ...base, planName: 'Professional' })).toBe('Professional');
    expect(toPlanLabel(base)).toBe('EDCONTROL_PRO');
    expect(toPlanLabel({ ...base, planId: undefined })).toBeUndefined();
  });
});

describe('needsMarketplaceCheckout', () => {
  const base: IFeatureSubscription = {
    feature: 'ED_CONTROL',
    state: 'inactive',
    isActive: false,
    isEnabled: false,
    subscriptionStatus: 'Inactive'
  };

  it('sends the user to pay only when nothing has been bought', () => {
    expect(needsMarketplaceCheckout(base)).toBe(true);
    // Paid for and switched off: switching on just brings it back.
    expect(needsMarketplaceCheckout({ ...base, state: 'active' })).toBe(false);
    // Already bought, on its way: a second checkout would be a second purchase.
    expect(needsMarketplaceCheckout({ ...base, state: 'pending' })).toBe(false);
    // Mended through Manage subscription, not by buying again.
    expect(needsMarketplaceCheckout({ ...base, state: 'suspended' })).toBe(false);
  });
});

describe('isMarketplaceReturn', () => {
  it('recognises the landing page marker', () => {
    expect(isMarketplaceReturn('?subscription=returned')).toBe(true);
    expect(isMarketplaceReturn('a=1&subscription=Returned')).toBe(true);
  });

  it('ignores anything else', () => {
    expect(isMarketplaceReturn('')).toBe(false);
    expect(isMarketplaceReturn('?subscription=active')).toBe(false);
    expect(isMarketplaceReturn('?subscription')).toBe(false);
    expect(isMarketplaceReturn('?returned=subscription')).toBe(false);
    expect(isMarketplaceReturn('?subscription=%E0%A4%A')).toBe(false);
  });
});

describe('readMarketplaceToken', () => {
  it('reads the token the landing page adds', () => {
    expect(MARKETPLACE_TOKEN_PARAMETER).toBe('marketplaceToken');
    expect(readMarketplaceToken('?subscription=returned&marketplaceToken=abc123')).toBe('abc123');
    expect(readMarketplaceToken('marketplaceToken=abc123')).toBe('abc123');
  });

  // Marketplace tokens are base64-like, so a `+` is data, not a space.
  it('keeps a plus sign and decodes percent escapes', () => {
    expect(readMarketplaceToken('?marketplaceToken=ab+cd%2Bef%3D%3D')).toBe('ab+cd+ef==');
  });

  it('answers undefined when there is no usable token', () => {
    expect(readMarketplaceToken('')).toBeUndefined();
    expect(readMarketplaceToken('?subscription=returned')).toBeUndefined();
    expect(readMarketplaceToken('?marketplaceToken=')).toBeUndefined();
    expect(readMarketplaceToken('?marketplaceToken')).toBeUndefined();
  });

  it('falls back to the raw value when it is not valid percent-encoding', () => {
    expect(readMarketplaceToken('?marketplaceToken=%E0%A4%A')).toBe('%E0%A4%A');
  });
});

describe('toActivatedFeatureSubscription', () => {
  it('maps an activation answer, reading the feature from the answer itself', () => {
    expect(toActivatedFeatureSubscription({ feature: 'KYP', isActive: true, subscriptionStatus: 'Subscribed' })).toEqual({
      feature: 'KYP',
      state: 'active',
      isActive: true,
      isEnabled: true,
      subscriptionStatus: 'Subscribed',
      planId: undefined,
      planName: undefined
    });
  });

  it('accepts the feature in lower case', () => {
    expect(toActivatedFeatureSubscription({ feature: 'tender', isActive: true })?.feature).toBe('TENDER');
  });

  it('rejects an unknown or missing feature, and a payload that is not an answer', () => {
    expect(toActivatedFeatureSubscription({ feature: 'TELEPORT', isActive: true })).toBeUndefined();
    expect(toActivatedFeatureSubscription({ feature: 'BIM', isActive: true })).toBeUndefined();
    expect(toActivatedFeatureSubscription({ isActive: true })).toBeUndefined();
    expect(toActivatedFeatureSubscription(null)).toBeUndefined();
    expect(toActivatedFeatureSubscription('KYP')).toBeUndefined();
    expect(toActivatedFeatureSubscription([{ feature: 'KYP' }])).toBeUndefined();
  });
});

describe('isNavigableUrl', () => {
  it('offers only absolute https links', () => {
    expect(isNavigableUrl('https://marketplace.microsoft.com/product/saas/contoso.xsproject')).toBe(true);
    expect(isNavigableUrl('')).toBe(false);
    expect(isNavigableUrl(undefined)).toBe(false);
    expect(isNavigableUrl('http://marketplace.microsoft.com')).toBe(false);
    expect(isNavigableUrl('/relative')).toBe(false);
    // Assembled so the linter's `no-script-url` does not mistake the test for the hazard.
    expect(isNavigableUrl(['javascript', 'alert(1)'].join(':'))).toBe(false);
  });
});

describe('offerUrlFor', () => {
  const OFFER: string = 'https://marketplace.microsoft.com/product/saas/contoso.xsproject';
  const KYP_PLAN: string = 'https://marketplace.microsoft.com/product/saas/contoso.xsproject?tab=PlansAndPrice&plan=kyp';

  it('sends a feature to its own link when one is configured', () => {
    const links = { marketplaceOfferUrl: OFFER, manageSubscriptionUrl: '', featureOfferUrls: { KYP: KYP_PLAN } };

    expect(offerUrlFor(links, 'KYP')).toBe(KYP_PLAN);
  });

  it('falls back to the offer link for a feature without one', () => {
    const links = { marketplaceOfferUrl: OFFER, manageSubscriptionUrl: '', featureOfferUrls: { KYP: KYP_PLAN } };

    expect(offerUrlFor(links, 'MINUTES')).toBe(OFFER);
    expect(offerUrlFor({ marketplaceOfferUrl: OFFER, manageSubscriptionUrl: '' }, 'KYP')).toBe(OFFER);
  });

  it('ignores a feature link that could not be opened, rather than sending the user nowhere', () => {
    const links = { marketplaceOfferUrl: OFFER, manageSubscriptionUrl: '', featureOfferUrls: { KYP: 'http://insecure', MINUTES: '' } };

    expect(offerUrlFor(links, 'KYP')).toBe(OFFER);
    expect(offerUrlFor(links, 'MINUTES')).toBe(OFFER);
  });

  it('reports "not configured" when neither link is set', () => {
    expect(offerUrlFor({ marketplaceOfferUrl: '', manageSubscriptionUrl: '' }, 'KYP')).toBe('');
  });
});
